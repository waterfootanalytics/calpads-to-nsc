# Contributing

For a bug report, include your operating system, the output of `prepare-nsc-graduates --version`, the command or menu choice, and the error message. Omit student information and account details. Use the private form for security reports. [Security reporting](SECURITY.md)

## Run the checks

From the source folder with a Python 3.11 or newer virtual environment active, run:

```sh
python -m pip install -e ".[dev]"
python -m ruff check src tests
python -m ruff format --check src tests
python -m unittest discover -s tests -v
```

Describe the change and how you verified it in your pull request. For changes to an input or output field, cite the applicable specification. Use type hints for CLI entry points, file readers and writers, and shared data contracts. Contributions use the [MIT License](LICENSE).

## Install the pinned runtime dependencies

```sh
python -m pip install --require-hashes --only-binary=:all: -r requirements.lock
python -m pip install --no-deps .
python -m pip check
prepare-nsc-graduates --version
```

## Release checks

1. For each reporting cycle, compare the CALPADS headers and completion codes and the NSC layout and submission instructions with the [official specifications](docs/file-formats.md). Update the source table with the date checked.
2. Run lint and the tests. Check the results for each supported operating system. When changing runtime dependencies, verify the wheel hashes against PyPI and update `requirements.lock`.
3. Build the wheel and source distribution. Install the wheel in a fresh environment, check the version and run the CLI. Inspect the files included in the release archive.
4. Update the [release notes](CHANGELOG.md) and package version together. Verify that the extracted source ZIP matches a clean checkout of the release commit. Tag that commit and attach the verified ZIP to the matching GitHub release.
