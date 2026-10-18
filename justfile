test:
    uv run pytest

ty:
    ty check src

ruff:
    ruff check src tests

lint: ruff ty

format:
    ruff format

cov:
    uv run pytest --cov --cov-report=html

outdated:
    uv pip list --outdated

lock:
    uv self update
    uv lock -U

sync:
    uv sync

doc: sync
    uv run sphinx-build -M singlehtml doc_src doc_src/_build
    cp -af doc_src/_build/singlehtml/* doc/

# local development
install-dev-odfdo:
    uv pip install "git+file://${PWD}/../odfdo/.git@devel"

dev-test: install-dev-odfdo test
