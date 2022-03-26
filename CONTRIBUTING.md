# Contributing to Fuzzy Table Extractor

## Setting up the environment 

The project uses [Poetry](https://python-poetry.org/) as a dependency manager, so to set up your environment after cloning the repository use the following command to install all dependencies:

```shell
poetry install
```

And this command to spawn a new shell with an active environment:

```shell
poetry shell
```
## Code and documentation style

Black is used as a code formatter, it is already listed as a dev dependency and will be installed when setting up the environment with Poetry.

Docstrings follow the [Google style](https://sphinxcontrib-napoleon.readthedocs.io/en/latest/example_google.html) and are used to generate code documentation by [Sphinx](https://www.sphinx-doc.org/en/master/).

## Tests

There are some tests written for the project, it still does not have 100% test coverage but it's increasing by the day. To submit a PR, make sure all the written tests pass and make sure to write new tests for the code you are adding.
To execute all the tests just run the following command:
```shell
pytest -v
```