"""
add_autoresize.py - Add column auto-resize to all files with Treeview
"""

import os
import re
from datetime import datetime

FILES_WITH_TREES = [
    "consumption.py",
    "donations.py",
    "end_users.py",
    "expiry_data.py",
    "item_families.py",
    "kits_Composition.py",
    "loans.py",
    "losses.py",
    "manage_items.py",
    "manage_parties.py",
    "manage_users.py",
    "order.py",
    "scenarios.py",
    "stock_availability.py",
    "stock_transactions.py",
    "dispatch_kit.py",
    "in_.py",
    "in_kit.py",
    "inv_kit.py",
    "out.py",
    "out_kit.py",
    "receive_kit.py",
    "reports.py",
    "standard_list.py",
    "stock_card.py",
    "stock_inv.py",
    "stock_summary.py",
]


def add_autoresize_to_file(filepath):
    """Add enable_column_auto_resize call to a file"""

    if not os.path.exists(filepath):
        return False

    print(f"\n📄 {filepath}")

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()

        # Check if already has auto-resize
        if "enable_column_auto_resize" in content:
            print(f"   ⏭️  Already has auto-resize")
            return False

        # Check if has Treeview
        if "ttk.Treeview" not in content and "self.tree" not in content:
            print(f"   ⏭️  No Treeview found")
            return False

        original = content

        # Step 1: Add import (if not already there)
        if "enable_column_auto_resize" not in content:
            # Find theme_config import line
            import_pattern = r"from theme_config import ([^\n]+)"
            match = re.search(import_pattern, content)

            if match:
                existing_imports = match.group(1).strip()
                if "enable_column_auto_resize" not in existing_imports:
                    new_imports = existing_imports + ", enable_column_auto_resize"
                    content = re.sub(
                        import_pattern,
                        f"from theme_config import {new_imports}",
                        content,
                        count=1,
                    )

        # Step 2: Find where to add the enable call
        # Look for _populate_tree or similar methods
        patterns = [
            (
                r"(def _populate_tree\(self\):.*?)((?=\n    def |\n\nclass |\Z))",
                "_populate_tree",
            ),
            (
                r"(def _render_table\(self\):.*?)((?=\n    def |\n\nclass |\Z))",
                "_render_table",
            ),
            (
                r"(def render\(self\):.*?self\.tree = ttk\.Treeview.*?)((?=\n    def |\n\nclass |\Z))",
                "render",
            ),
        ]

        added = False
        for pattern, method_name in patterns:
            match = re.search(pattern, content, re.DOTALL)
            if match and not added:
                method_content = match.group(1)

                # Check if already has the call
                if "enable_column_auto_resize" in method_content:
                    continue

                # Find the end of the method (before next def or class)
                # Add the call at the end
                insertion = "\n        # Enable double-click column auto-resize\n        enable_column_auto_resize(self.tree)\n"

                # Insert before the closing of the method
                end_match = match.group(2)
                content = content.replace(
                    match.group(0), match.group(1) + insertion + end_match
                )
                added = True
                break

        if not added:
            print(f"   ⚠️  Could not find suitable location to add auto-resize")
            return False

        if content == original:
            print(f"   ⏭️  No changes made")
            return False

        # Backup
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = f"{filepath}.autoresize_backup_{timestamp}"
        with open(backup, "w", encoding="utf-8") as f:
            f.write(original)

        # Save
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)

        print(f"   ✅ Added auto-resize support")
        print(f"   💾 Backup: {backup}")
        return True

    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def main():
    print("=" * 70)
    print("🎯 ADD COLUMN AUTO-RESIZE TO ALL TREEVIEWS")
    print("=" * 70)
    print("\nThis will add double-click column auto-resize to all files.")
    print("\nFeatures:")
    print("  • Double-click column header → auto-fit to content")
    print("  • Measures both header and data")
    print("  • Adds one line of code per file")
    print("\n" + "=" * 70)

    existing = [f for f in FILES_WITH_TREES if os.path.exists(f)]

    print(f"\n📋 Found {len(existing)} files")

    response = input("\n❓ Proceed? (yes/no): ").strip().lower()
    if response not in ["yes", "y"]:
        print("\n❌ Cancelled")
        return

    print("\n" + "=" * 70)
    print("⚙️  PROCESSING...")
    print("=" * 70)

    updated = 0
    skipped = 0

    for filepath in existing:
        if add_autoresize_to_file(filepath):
            updated += 1
        else:
            skipped += 1

    print("\n" + "=" * 70)
    print("✅ COMPLETE!")
    print("=" * 70)
    print(f"\n📊 Summary:")
    print(f"   Files processed: {len(existing)}")
    print(f"   Files updated: {updated}")
    print(f"   Files skipped: {skipped}")

    if updated > 0:
        print("\n📝 Test: python main.py")
        print("   Double-click any column header to auto-resize!")


if __name__ == "__main__":
    main()
