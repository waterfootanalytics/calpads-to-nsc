import os
import tempfile
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path


@contextmanager
def private_output_path(directory: Path, suffix: str = "") -> Iterator[Path]:
    descriptor, filename = tempfile.mkstemp(dir=directory, suffix=suffix)
    os.close(descriptor)
    path = Path(filename)
    try:
        yield path
    finally:
        path.unlink(missing_ok=True)


def publish_new(temp: Path, destination: Path) -> None:
    if os.name == "nt":
        os.rename(temp, destination)
        return
    try:
        os.link(temp, destination)
        return
    except FileExistsError:
        raise
    except OSError:
        pass
    descriptor = os.open(destination, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    try:
        reservation = os.fstat(descriptor)
        try:
            os.replace(temp, destination)
        except OSError:
            if destination.exists() and os.path.samestat(reservation, destination.stat()):
                destination.unlink()
            raise
    finally:
        os.close(descriptor)


@contextmanager
def private_update_lock(path: Path) -> Iterator[None]:
    marker = path.with_name("." + path.name + ".editing")
    deadline = time.monotonic() + 3
    while True:
        try:
            marker.mkdir(mode=0o700)
            break
        except FileExistsError:
            if time.monotonic() >= deadline:
                raise BlockingIOError(f"Another session is editing {path.name}.") from None
            time.sleep(0.05)
    try:
        yield
    finally:
        marker.rmdir()
