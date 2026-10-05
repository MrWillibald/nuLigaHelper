"""Run dependency-free DOM-level assignment regressions."""

import shutil
import subprocess

import helpers as h


def test_released_person_keeps_server_membership_metadata_and_sort_position():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run(
        [node, "test/js_assignment_options.mjs"],
        cwd=h.PROJECT_DIR,
        text=True,
        capture_output=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout


def test_rejected_claim_does_not_change_displayed_progress():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run(
        [node, "test/js_schedule_progress_events.mjs"],
        cwd=h.PROJECT_DIR,
        text=True,
        capture_output=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout


def test_candidate_loading_handles_open_cards_failures_and_slot_metadata():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run(
        [node, "test/js_candidate_loading.mjs"],
        cwd=h.PROJECT_DIR,
        text=True,
        capture_output=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout


def test_expanding_a_card_requests_its_candidates_and_offers_retry():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run(
        [node, "test/js_candidate_toggle.mjs"],
        cwd=h.PROJECT_DIR,
        text=True,
        capture_output=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout


def test_task_help_supports_pointer_keyboard_touch_without_assignment_actions():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run(
        [node, "test/js_task_descriptions.mjs"],
        cwd=h.PROJECT_DIR,
        text=True,
        capture_output=True,
        check=False,
        timeout=30,
    )
    assert result.returncode == 0, result.stderr or result.stdout


if __name__ == "__main__":
    h.run_all(dict(globals()))
