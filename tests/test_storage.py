import errno
import os
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier
from unittest.mock import patch

from calpads_to_nsc.storage import private_output_path, private_update_lock, publish_new


class OutputStorageTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.destination = self.folder / "output.txt"

    def save(self, content):
        with private_output_path(self.folder) as temp:
            temp.write_bytes(content)
            publish_new(temp, self.destination)

    def test_no_hard_links_and_concurrent_writers(self):
        start = Barrier(2)

        def attempt(content):
            start.wait(timeout=5)
            try:
                self.save(content)
            except FileExistsError:
                return False
            return True

        payloads = [b"a" * 15000, b"b" * 25000]
        with patch("os.link", side_effect=OSError(errno.ENOTSUP, "unsupported")), ThreadPoolExecutor(2) as pool:
            results = list(pool.map(attempt, payloads))
        self.assertEqual(sum(results), 1)
        self.assertEqual(self.destination.read_bytes(), payloads[results.index(True)])
        self.assertEqual(list(self.folder.iterdir()), [self.destination])
        if os.name == "posix":
            self.assertEqual(self.destination.stat().st_mode & 0o777, 0o600)

    @unittest.skipUnless(os.name == "posix", "POSIX reservation fallback")
    def test_failed_publication_removes_reservation_and_can_retry(self):
        with patch("os.link", side_effect=OSError(errno.ENOTSUP, "unsupported")):
            with patch("os.replace", side_effect=PermissionError), self.assertRaises(PermissionError):
                self.save(b"complete")
            self.assertEqual(list(self.folder.iterdir()), [])
            self.save(b"retry")
        self.assertEqual(self.destination.read_bytes(), b"retry")

    def test_failed_decision_update_releases_local_edit_lock(self):
        with self.assertRaises(ValueError), private_update_lock(self.destination):
            raise ValueError("failed update")
        with private_update_lock(self.destination):
            with patch("time.monotonic", side_effect=[0.0, 4.0]), self.assertRaises(BlockingIOError):
                with private_update_lock(self.destination):
                    self.fail("another writer acquired the lock")
        self.assertEqual(list(self.folder.iterdir()), [])
