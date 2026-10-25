from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import ios_xe_upgrade_manager as manager


class SharedStorageLayoutTests(unittest.TestCase):
    def test_images_and_backups_share_one_netupgrade_root(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as temp:
            root = Path(temp) / "chosen" / "NetUpgrade"
            settings = {"StoragePath": str(root), "ImagesPath": None, "BackupsPath": None}
            with patch.object(manager, "_storage_setting", side_effect=lambda key: settings.get(key)), \
                    patch.object(manager, "this_program_path", return_value=Path(temp) / "launcher.exe"):
                image_path = manager.images_dir()
                backup_path = manager.sites_root()

            self.assertEqual(image_path, root / manager.IMAGES_DIRNAME)
            self.assertEqual(backup_path, root / manager.SITES_ROOT_DIRNAME)
            self.assertEqual(image_path.parent, backup_path.parent)
            self.assertTrue(image_path.is_dir())
            self.assertTrue(backup_path.is_dir())

    def test_migrates_both_previous_installer_folders(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as temp:
            temp_path = Path(temp)
            root = temp_path / "new" / "NetUpgrade"
            old_images = temp_path / "old-images"
            old_backups = temp_path / "old-backups"
            old_images.mkdir()
            (old_images / "cat9k.bin").write_bytes(b"image")
            (old_images / "cat9k.bin.md5").write_text("0123456789abcdef0123456789abcdef", encoding="ascii")
            (old_backups / "Site A" / "Switch Backups" / "sw1").mkdir(parents=True)
            (old_backups / "Site A" / "Switch Backups" / "sw1" / "backup.txt").write_text("config", encoding="utf-8")
            settings = {"StoragePath": str(root), "ImagesPath": str(old_images), "BackupsPath": str(old_backups)}

            with patch.object(manager, "_storage_setting", side_effect=lambda key: settings.get(key)), \
                    patch.object(manager, "this_program_path", return_value=temp_path / "launcher.exe"):
                manager.images_dir()
                manager.sites_root()

            self.assertEqual((root / manager.IMAGES_DIRNAME / "cat9k.bin").read_bytes(), b"image")
            self.assertEqual((root / manager.IMAGES_DIRNAME / "cat9k.bin.md5").read_text(encoding="ascii"),
                             "0123456789abcdef0123456789abcdef")
            self.assertEqual((root / manager.SITES_ROOT_DIRNAME / "Site A" / "Switch Backups" / "sw1" / "backup.txt")
                             .read_text(encoding="utf-8"), "config")

    def test_does_not_overwrite_existing_files_during_migration(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as temp:
            temp_path = Path(temp)
            root = temp_path / "new" / "NetUpgrade"
            old_images = temp_path / "old-images"
            old_images.mkdir()
            (old_images / "cat9k.bin").write_bytes(b"old")
            destination = root / manager.IMAGES_DIRNAME
            destination.mkdir(parents=True)
            (destination / "cat9k.bin").write_bytes(b"new")
            settings = {"StoragePath": str(root), "ImagesPath": str(old_images), "BackupsPath": None}

            with patch.object(manager, "_storage_setting", side_effect=lambda key: settings.get(key)), \
                    patch.object(manager, "this_program_path", return_value=temp_path / "launcher.exe"):
                manager.images_dir()

            self.assertEqual((destination / "cat9k.bin").read_bytes(), b"new")

    def test_default_root_uses_netupgrade_folder_next_to_program(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as temp:
            temp_path = Path(temp)
            settings = {"StoragePath": None, "ImagesPath": None, "BackupsPath": None}
            with patch.object(manager, "_storage_setting", side_effect=lambda key: settings.get(key)), \
                    patch.object(manager, "this_program_path", return_value=temp_path / "launcher.exe"):
                self.assertEqual(manager.common_storage_root(), temp_path / manager.STORAGE_DIRNAME)


if __name__ == "__main__":
    unittest.main()
