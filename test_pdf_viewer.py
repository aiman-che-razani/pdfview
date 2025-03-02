import os
import pytest
import base64
import tempfile  # <-- Make sure this import is here
from unittest.mock import patch
from pdf_viewer import list_pdfs, display_pdf


@pytest.fixture
def setup_test_pdfs(tmp_path):
    """Creates a temporary directory with sample PDF files for testing."""
    pdf1 = tmp_path / "file1.pdf"
    pdf2 = tmp_path / "file2.pdf"
    non_pdf = tmp_path / "not_a_pdf.txt"

    pdf1.touch()
    pdf2.touch()
    non_pdf.touch()

    return tmp_path


def test_list_pdfs(setup_test_pdfs):
    """Tests listing PDF files in a directory."""
    pdf_files = list_pdfs(setup_test_pdfs)
    assert sorted(pdf_files) == ["file1.pdf", "file2.pdf"]


def test_display_pdf():
    """Tests that the PDF display function reads
    files correctly without testing UI."""
    # Create a temporary PDF file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(b"Test PDF content")
        temp_pdf_path = temp_pdf.name

    try:
        # Read the file manually for comparison
        with open(temp_pdf_path, "rb") as pdf_file:
            expected_base64 = base64.b64encode(pdf_file.read()).decode("utf-8")

        with patch("streamlit.markdown") as mock_markdown:
            display_pdf(temp_pdf_path)
            mock_markdown.assert_called_once()
            assert expected_base64 in mock_markdown.call_args[0][0]
            # Ensure the PDF was encoded correctly
    finally:
        os.remove(temp_pdf_path)  # Clean up the test file
