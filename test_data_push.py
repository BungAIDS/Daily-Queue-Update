import json
import tempfile
from pathlib import Path
from unittest import mock

import data_push


def _read_published(files):
    return {file.name: json.loads(file.data.decode("utf-8"))
            for file in files if file.data is not None}


def test_small_and_non_job_json_keep_the_original_file(tmp_path):
    small = tmp_path / "small.json"
    small.write_text('{"hello": "world"}\n', encoding="utf-8")
    non_job = tmp_path / "settings.json"
    non_job.write_text(json.dumps({"text": "x" * 300}), encoding="utf-8")

    with mock.patch.object(data_push, "MAX_JSON_BYTES", 100):
        published = data_push.prepare_publish_files([small, non_job])

    assert [file.name for file in published] == ["small.json", "settings.json"]
    assert all(file.path == path for file, path in zip(published, [small, non_job]))


def test_large_direct_job_store_uses_deterministic_range_shards_and_manifest(tmp_path):
    source = tmp_path / "line_items.json"
    source.write_text(json.dumps({str(job): {"value": "x" * 25}
                                  for job in [421002, 421001, 421004, 421003]}, indent=2),
                      encoding="utf-8")

    with mock.patch.object(data_push, "MAX_JSON_BYTES", 170), \
            mock.patch.object(data_push, "JOB_RANGE_SIZE", 2):
        first = data_push.prepare_publish_files([source])
        second = data_push.prepare_publish_files([source])

    assert [file.name for file in first] == [file.name for file in second]
    assert [file.data for file in first] == [file.data for file in second]
    assert first[-1].name == "line_items.manifest.json"
    assert all(file.name != "line_items.json" for file in first)
    manifest = _read_published(first)["line_items.manifest.json"]
    assert manifest["format"] == "order-data-shards-v1"
    assert manifest["wrapper"] is None
    assert manifest["job_count"] == 4
    assert [(shard["first_job"], shard["last_job"]) for shard in manifest["shards"]] == [
        (421001, 421001), (421002, 421003), (421004, 421004)
    ]

    combined = {}
    for file in first[:-1]:
        combined.update(json.loads(file.data.decode("utf-8")))
    assert list(combined) == ["421001", "421002", "421003", "421004"]


def test_large_wrapped_job_store_preserves_wrapper_key(tmp_path):
    source = tmp_path / "live_master.json"
    source.write_text(json.dumps({
        "version": 2,
        "orders": {job: {"job": job, "value": "y" * 20}
                   for job in ["421001", "421002A", "421003"]},
    }, indent=2), encoding="utf-8")

    with mock.patch.object(data_push, "MAX_JSON_BYTES", 150), \
            mock.patch.object(data_push, "JOB_RANGE_SIZE", 2):
        published = data_push.prepare_publish_files([source])

    assert published[-1].name == "live_master.manifest.json"
    for file in published[:-1]:
        payload = json.loads(file.data.decode("utf-8"))
        assert set(payload) == {"orders"}
        assert list(payload["orders"])
    manifest = _read_published(published)["live_master.manifest.json"]
    assert manifest["metadata"] == {"version": 2}


def test_build_snapshot_commit_remains_orphan_and_includes_shards(tmp_path):
    source = tmp_path / "store.json"
    source.write_text(json.dumps({str(job): {"value": "z" * 40}
                                  for job in [421001, 421002]}), encoding="utf-8")

    with mock.patch.object(data_push, "MAX_JSON_BYTES", 100), \
            mock.patch.object(data_push, "JOB_RANGE_SIZE", 1):
        commit = data_push.build_snapshot_commit([source], "test snapshot")

    assert commit
    parents = data_push._git(["show", "-s", "--format=%P", commit])
    tree = data_push._git(["ls-tree", "--name-only", commit])
    assert parents.returncode == 0 and not parents.stdout.strip()
    assert "store.manifest.json" in tree.stdout
    assert "store.jobs-421001-421001.json" in tree.stdout


def test_remote_verification_failure_is_not_reported_as_success():
    fake_file = data_push._PublishedFile(name="tiny.json", data=b"{}\n")
    completed = data_push.subprocess.CompletedProcess
    with mock.patch.object(data_push, "data_files", return_value=[Path("tiny.json")]), \
            mock.patch.object(data_push, "prepare_publish_files", return_value=[fake_file]), \
            mock.patch.object(data_push, "_build_snapshot_commit_prepared", return_value="abc123"), \
            mock.patch.object(data_push, "_git", side_effect=[
                completed([], 0, "", ""),
                completed([], 1, "", "network unavailable"),
            ]):
        assert data_push.push_data("order-data-test") is False


def main():
    tests = [
        test_small_and_non_job_json_keep_the_original_file,
        test_large_direct_job_store_uses_deterministic_range_shards_and_manifest,
        test_large_wrapped_job_store_preserves_wrapper_key,
        test_build_snapshot_commit_remains_orphan_and_includes_shards,
    ]
    with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
        root = Path(temp_dir)
        for index, test in enumerate(tests):
            case_dir = root / str(index)
            case_dir.mkdir()
            test(case_dir)
    test_remote_verification_failure_is_not_reported_as_success()
    print("All data_push tests passed.")


if __name__ == "__main__":
    main()
