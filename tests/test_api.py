"""API integration tests."""

import pytest
from fastapi.testclient import TestClient

from docx_formatter.api.main import app


@pytest.fixture
def client():
    """Test client fixture."""
    return TestClient(app)


def test_health_check(client):
    """Test health endpoint."""
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"


def test_format_upload_invalid_extension(client):
    """Test upload with invalid file extension."""
    response = client.post(
        "/api/v1/format/template-upload",
        files={
            "template": ("bad.txt", b"not docx", "text/plain"),
            "content": ("content.txt", b"not docx", "text/plain"),
        },
    )
    assert response.status_code == 422


@pytest.mark.parametrize(
    "template_name,content_name",
    [
        ("template_offer.docx", "content_offer.docx"),
        ("template_cv.docx", "content_cv.docx"),
    ],
)
def test_format_upload_success(client, template_name, content_name):
    """Test successful formatting with fixtures."""
    fixtures_dir = __file__.replace("test_api.py", "fixtures/")

    with open(f"{fixtures_dir}{template_name}", "rb") as t, open(
        f"{fixtures_dir}{content_name}", "rb"
    ) as c:
        response = client.post(
            "/api/v1/format/template-upload",
            files={
                "template": (template_name, t, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                "content": (content_name, c, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            },
            data={"output_filename": "output.docx"},
        )

    assert response.status_code == 200, response.text
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    assert response.headers["content-disposition"].startswith("attachment")
    assert len(response.content) > 1000  # DOCX is a ZIP, should be > 1KB


def test_format_upload_custom_filename(client):
    """Test with custom output filename."""
    fixtures_dir = __file__.replace("test_api.py", "fixtures/")

    with open(f"{fixtures_dir}template_offer.docx", "rb") as t, open(
        f"{fixtures_dir}content_offer.docx", "rb"
    ) as c:
        response = client.post(
            "/api/v1/format/template-upload",
            files={
                "template": ("template_offer.docx", t, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                "content": ("content_offer.docx", c, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            },
            data={"output_filename": "my_offer.docx"},
        )

    assert response.status_code == 200
    assert "my_offer.docx" in response.headers["content-disposition"]


def test_format_upload_debug(client):
    """Test debug endpoint returns matching logs."""
    fixtures_dir = __file__.replace("test_api.py", "fixtures/")

    with open(f"{fixtures_dir}template_offer.docx", "rb") as t, open(
        f"{fixtures_dir}content_offer.docx", "rb"
    ) as c:
        response = client.post(
            "/api/v1/format/template-upload/debug",
            files={
                "template": ("template_offer.docx", t, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                "content": ("content_offer.docx", c, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            },
        )

    assert response.status_code == 200, response.text
    data = response.json()
    assert data["success"] is True
    assert data["template_styles_found"] > 0
    assert data["content_paragraphs"] > 0
    assert "total_matches" in data
    assert isinstance(data["matches"], list)
    assert "llm_available" in data
    assert "llm_used" in data
    assert "unmatched" in data


def test_format_upload_debug_invalid_file(client):
    """Test debug endpoint with invalid files."""
    response = client.post(
        "/api/v1/format/template-upload/debug",
        files={
            "template": ("bad.txt", b"not docx", "text/plain"),
            "content": ("content.txt", b"not docx", "text/plain"),
        },
    )
    assert response.status_code == 422
