test:
    uv run pytest

lint:
    ruff check src tests
    mypy

lock:
    uv lock -U

sync:
    uv sync

# local development
install-dev-odfdo:
    uv pip install "git+file://${PWD}/../odfdo/.git@devel"

dev-test: install-dev-odfdo test
