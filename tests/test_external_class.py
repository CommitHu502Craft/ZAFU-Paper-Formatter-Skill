from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from formatter_backends.external_class import CLASS_SHA256, install_external_class


class ExternalClassTests(unittest.TestCase):
    def test_download_verified_once_then_reused_offline(self) -> None:
        data = b"test external class bytes"
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "cache"
            project = Path(directory) / "project"
            project.mkdir()
            response = MagicMock()
            response.__enter__.return_value.read.return_value = data
            with patch("formatter_backends.external_class.CLASS_SHA256", hashlib.sha256(data).hexdigest()), patch("formatter_backends.external_class.urllib.request.urlopen", return_value=response) as download:
                first = install_external_class(project, cache)
                second = install_external_class(project, cache)
                self.assertEqual(download.call_count, 1)
            self.assertEqual(first, second)
            self.assertEqual((project / "ZafuThesis.cls").read_bytes(), data)
            self.assertEqual(json.loads((project / "UPSTREAM.json").read_text())["modified"], False)
            self.assertEqual(list(cache.glob("*.tmp")), [])

    def test_corrupt_cache_is_rejected_without_network_or_publication(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "cache"
            cache.mkdir()
            (cache / "ZafuThesis.cls").write_bytes(b"corrupt")
            project = Path(directory) / "project"
            project.mkdir()
            with patch("formatter_backends.external_class.urllib.request.urlopen") as download:
                with self.assertRaisesRegex(ValueError, "integrity"):
                    install_external_class(project, cache)
                download.assert_not_called()
            self.assertFalse((project / "ZafuThesis.cls").exists())

    def test_unverified_download_is_never_cached(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "cache"
            project = Path(directory) / "project"
            project.mkdir()
            response = MagicMock()
            response.__enter__.return_value.read.return_value = b"invalid remote bytes"
            with patch("formatter_backends.external_class.urllib.request.urlopen", return_value=response):
                with self.assertRaisesRegex(ValueError, "integrity"):
                    install_external_class(project, cache)
            self.assertFalse((cache / "ZafuThesis.cls").exists())

    def test_local_upstream_copy_is_byte_exact_when_available(self) -> None:
        cached = ROOT / ".local/dependencies/zafu-template/ZafuThesis.cls"
        if not cached.exists():
            self.skipTest("Local upstream cache not provisioned")
        self.assertEqual(hashlib.sha256(cached.read_bytes()).hexdigest(), CLASS_SHA256)


if __name__ == "__main__":
    unittest.main()
