"""Development entry point for Reportus."""

import argparse

from core.agent import ReportusAgent


def main() -> None:
    """Run a lightweight development command until the desktop shell lands."""

    parser = argparse.ArgumentParser(prog="reportus")
    parser.add_argument(
        "--list-skills",
        action="store_true",
        help="list validated report functions",
    )
    args = parser.parse_args()

    agent = ReportusAgent()
    if args.list_skills:
        for skill_id, label in agent.available_reports():
            print(f"{skill_id}\t{label}")
        if agent.registry.issues:
            for issue in agent.registry.issues:
                print(f"INVALID\t{issue.skill_id}\t{issue.message}")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
