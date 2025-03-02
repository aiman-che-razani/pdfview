import os
import pytest
from app import get_second_level_folders

@pytest.fixture
def setup_test_folders(tmp_path):
    """Creates a temporary directory with nested subfolders for testing."""
    root = tmp_path / "source_files"
    root.mkdir()

    first_layer = root / "Folder_A"
    first_layer.mkdir()
    (root / "Folder_B").mkdir()

    second_layer_a = first_layer / "Subfolder_A1"
    second_layer_a.mkdir()
    (first_layer / "Subfolder_A2").mkdir()

    second_layer_b = root / "Folder_B" / "Subfolder_B1"
    second_layer_b.mkdir()

    return root

def test_get_second_level_folders(setup_test_folders):
    """Test function for retrieving second-level folders."""
    root_folder = str(setup_test_folders)
    first_level_folders, second_level_folders = get_second_level_folders(root_folder)

    # Check first-level folder detection
    assert sorted(first_level_folders) == ["Folder_A", "Folder_B"]

    # Check second-level folder detection
    assert "Folder_A" in second_level_folders
    assert sorted(second_level_folders["Folder_A"]) == ["Subfolder_A1", "Subfolder_A2"]

    assert "Folder_B" in second_level_folders
    assert sorted(second_level_folders["Folder_B"]) == ["Subfolder_B1"]

    # Test empty case
    empty_root = os.path.join(root_folder, "Empty_Folder")
    os.mkdir(empty_root)
    first_level_empty, second_level_empty = get_second_level_folders(empty_root)
    assert first_level_empty == []
    assert second_level_empty == {}
