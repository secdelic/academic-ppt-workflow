from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.project_cache import (  # noqa: E402
    CachePolicyError,
    ProjectCache,
    resolve_project_cache_root,
    validate_cache_root,
)
from academic_ppt.project_interface import ProjectPaths  # noqa: E402


class ExternalWorkspaceCacheContractTests(unittest.TestCase):
    fixture_provenance = "SYNTHETIC"

    def roots(self, temporary: str) -> tuple[Path, Path, Path]:
        base = Path(temporary)
        app = base / "application"
        workspace = base / "workspace"
        project = workspace / "projects" / "P1"
        app.mkdir(parents=True)
        project.mkdir(parents=True)
        return app, workspace, project

    def test_external_workspace_default_cache(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            actual = resolve_project_cache_root(app, project_root=project)
            self.assertEqual(actual, (project / "cache").resolve())
            self.assertFalse(actual.is_relative_to((app / ".cache").resolve()))

    def test_ppt_cache_home_is_shared_parent_with_opaque_project_child(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, workspace, project = self.roots(temporary)
            shared = workspace / "shared_ppt_cache"
            with patch.dict(os.environ, {"PPT_CACHE_HOME": str(shared)}):
                cache = ProjectCache.from_identity(
                    app,
                    "SYNTHETIC_PRIVATE_PROJECT",
                    project_root=project,
                    production=True,
                )
            self.assertEqual(cache.cache_root, shared.resolve())
            self.assertTrue(cache.project_dir.is_relative_to(shared.resolve()))
            self.assertRegex(cache.project_dir.name, r"^PRJ-[A-F0-9]{32}$")
            self.assertNotIn("SYNTHETIC_PRIVATE_PROJECT", str(cache.project_dir))

    def test_repository_cache_is_dev_test_only(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            expected = (app / ".cache" / "project_state").resolve()
            self.assertEqual(validate_cache_root(app), expected)
            with self.assertRaises(CachePolicyError):
                resolve_project_cache_root(
                    app,
                    project_root=project,
                    explicit_cache_root=expected,
                    production=True,
                )

    def test_cross_project_cache_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, workspace, project_a = self.roots(temporary)
            project_b = workspace / "projects" / "P2"
            project_b.mkdir(parents=True)
            with self.assertRaises(CachePolicyError):
                resolve_project_cache_root(
                    app,
                    project_root=project_a,
                    explicit_cache_root=project_b / "cache",
                )

    def test_shared_cache_uses_distinct_opaque_key_for_each_project(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, workspace, project_a = self.roots(temporary)
            project_b = workspace / "projects" / "P2"
            project_b.mkdir(parents=True)
            shared = workspace / "shared_ppt_cache"
            first = ProjectCache.from_identity(
                app,
                "SAME_DISPLAY_NAME",
                project_root=project_a,
                cache_home=shared,
                production=True,
            )
            second = ProjectCache.from_identity(
                app,
                "SAME_DISPLAY_NAME",
                project_root=project_b,
                cache_home=shared,
                production=True,
            )
            self.assertNotEqual(first.project_key, second.project_key)
            self.assertFalse(first.project_dir.is_relative_to(second.project_dir))

            with self.assertRaises(CachePolicyError):
                ProjectCache(
                    app,
                    second.project_key,
                    cache_root=shared,
                    project_root=project_a,
                    cache_home=shared,
                    production=True,
                )

    def test_parent_traversal_is_rejected_before_resolve(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            with self.assertRaises(CachePolicyError):
                resolve_project_cache_root(
                    app,
                    project_root=project,
                    explicit_cache_root=project / "cache" / ".." / ".." / "other",
                )

    def test_unconfigured_external_directory_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, workspace, project = self.roots(temporary)
            unrelated = workspace / "unrelated"
            with self.assertRaises(CachePolicyError):
                resolve_project_cache_root(
                    app,
                    project_root=project,
                    explicit_cache_root=unrelated,
                )

    def test_configured_cache_home_inside_repository_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            with self.assertRaises(CachePolicyError):
                resolve_project_cache_root(
                    app,
                    project_root=project,
                    cache_home=app / "src" / "cache",
                )

    def test_clinical_identity_is_not_written_into_cache_path(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            patient_like_identity = "BED-21-PATIENT-000123"
            cache = ProjectCache.from_identity(
                app,
                patient_like_identity,
                project_root=project,
                production=True,
                clinical_privacy_mode=True,
            )
            self.assertRegex(cache.project_dir.name, r"^PRJ-[A-F0-9]{32}$")
            self.assertNotIn(patient_like_identity, str(cache.project_dir))

    def test_project_paths_define_route_independent_default_cache(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            _, _, project = self.roots(temporary)
            for route in ("generate", "enhance", "template-fill"):
                with self.subTest(route=route):
                    paths = ProjectPaths.from_root(project)
                    self.assertEqual(paths.cache_root, (project / "cache").resolve())

    def test_second_device_external_workspace_cache_regression(self) -> None:
        """SECOND_DEVICE_EXTERNAL_WORKSPACE_CACHE regression fixture."""

        with tempfile.TemporaryDirectory() as temporary:
            app, _, project = self.roots(temporary)
            resolved = resolve_project_cache_root(app, project_root=project)
            cache = ProjectCache.from_identity(
                app,
                "SECOND_DEVICE_EXTERNAL_WORKSPACE_CACHE",
                cache_root=resolved,
                project_root=project,
                production=True,
            )
            self.assertEqual(cache.cache_root, (project / "cache").resolve())
            self.assertTrue(cache.project_dir.is_dir())
            self.assertFalse(cache.project_dir.is_relative_to(app.resolve()))


if __name__ == "__main__":
    unittest.main()
