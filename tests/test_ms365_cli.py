import argparse
import unittest
from unittest.mock import patch

import ms365_cli


class MS365CliSafetyTests(unittest.TestCase):
    def test_guarded_write_dry_runs_by_default(self):
        args = argparse.Namespace(confirm=False, dry_run=False)

        with patch.object(ms365_cli, "call_mcp") as call_mcp:
            result = ms365_cli.guarded_write(
                args,
                "send email",
                "send-mail",
                {"body": {"subject": "Hello"}},
            )

        call_mcp.assert_not_called()
        self.assertEqual(result["dry_run"], True)
        self.assertEqual(result["mcp_method"], "send-mail")
        self.assertIn("--confirm", result["message"])

    def test_guarded_write_calls_mcp_only_when_confirmed(self):
        args = argparse.Namespace(confirm=True, dry_run=False)

        with patch.object(ms365_cli, "call_mcp", return_value={"ok": True}) as call_mcp:
            result = ms365_cli.guarded_write(
                args,
                "create task",
                "create-todo-task",
                {"todoTaskListId": "list-1"},
            )

        call_mcp.assert_called_once_with("create-todo-task", {"todoTaskListId": "list-1"})
        self.assertEqual(result, {"ok": True})

    def test_dry_run_overrides_confirm(self):
        args = argparse.Namespace(confirm=True, dry_run=True)

        with patch.object(ms365_cli, "call_mcp") as call_mcp:
            result = ms365_cli.guarded_write(
                args,
                "create calendar event",
                "create-calendar-event",
                {"body": {"subject": "Planning"}},
            )

        call_mcp.assert_not_called()
        self.assertEqual(result["dry_run"], True)

    def test_mail_send_dry_runs_without_calling_mcp(self):
        args = argparse.Namespace(
            to="user@example.com",
            subject="Project update",
            body="Hello",
            confirm=False,
            dry_run=False,
        )

        with (
            patch.object(ms365_cli, "call_mcp") as call_mcp,
            patch.object(ms365_cli, "format_output") as format_output,
        ):
            ms365_cli.cmd_mail_send(args)

        call_mcp.assert_not_called()
        output = format_output.call_args.args[0]
        self.assertEqual(output["dry_run"], True)
        self.assertEqual(output["mcp_method"], "send-mail")
        self.assertEqual(
            output["arguments"]["body"]["message"]["toRecipients"][0]["emailAddress"]["address"],
            "user@example.com",
        )

    def test_calendar_create_confirmed_calls_mcp(self):
        args = argparse.Namespace(
            subject="Planning",
            start="2026-05-16T10:00:00",
            end="2026-05-16T10:30:00",
            body="Agenda",
            timezone="America/New_York",
            confirm=True,
            dry_run=False,
        )

        with (
            patch.object(ms365_cli, "call_mcp", return_value={"id": "event-1"}) as call_mcp,
            patch.object(ms365_cli, "format_output") as format_output,
        ):
            ms365_cli.cmd_calendar_create(args)

        call_mcp.assert_called_once()
        method, params = call_mcp.call_args.args
        self.assertEqual(method, "create-calendar-event")
        self.assertEqual(params["body"]["subject"], "Planning")
        self.assertEqual(params["body"]["start"]["timeZone"], "America/New_York")
        format_output.assert_called_once_with({"id": "event-1"})

    def test_task_create_dry_runs_with_due_payload(self):
        args = argparse.Namespace(
            list_id="list-1",
            title="Review budget",
            due="2026-05-20",
            confirm=False,
            dry_run=False,
        )

        with (
            patch.object(ms365_cli, "call_mcp") as call_mcp,
            patch.object(ms365_cli, "format_output") as format_output,
        ):
            ms365_cli.cmd_tasks_create(args)

        call_mcp.assert_not_called()
        output = format_output.call_args.args[0]
        self.assertEqual(output["mcp_method"], "create-todo-task")
        self.assertEqual(output["arguments"]["todoTaskListId"], "list-1")
        self.assertEqual(output["arguments"]["body"]["dueDateTime"]["dateTime"], "2026-05-20")


if __name__ == "__main__":
    unittest.main()
