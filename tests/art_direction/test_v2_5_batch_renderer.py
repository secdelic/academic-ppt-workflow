from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]


class BatchRendererContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.runner = (ROOT / "scripts" / "v2_5" / "run_v2_5.py").read_text(encoding="utf-8")
        cls.worker = (ROOT / "scripts" / "v2_5" / "render_v2_5_worker.mjs").read_text(encoding="utf-8")

    def test_worker_is_stdin_owned_and_has_no_network_service(self):
        self.assertIn("readline.createInterface({ input: process.stdin", self.worker)
        self.assertNotIn("createServer", self.worker)
        self.assertNotIn("listen(", self.worker)
        self.assertNotIn("fetch(", self.worker)

    def test_each_job_gets_an_isolated_renderer_module_instance(self):
        self.assertIn("?batch_job=", self.worker)
        self.assertIn("job.graph_path", self.worker)
        self.assertIn("job.output_path", self.worker)

    def test_runner_fails_closed_on_worker_or_output_failure(self):
        self.assertIn("if not result or not result.get(\"ok\")", self.runner)
        self.assertIn("not pptx.is_file() or pptx.stat().st_size == 0", self.runner)
        self.assertIn("renderer_worker.close()", self.runner)


if __name__ == "__main__":
    unittest.main()
