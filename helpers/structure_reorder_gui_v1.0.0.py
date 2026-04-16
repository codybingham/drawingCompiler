import os
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import filedialog, messagebox, ttk

import pandas as pd

APP_VERSION = "1.0.2"


@dataclass
class StructureNode:
    level: str
    description: str
    part_number: str
    children: list["StructureNode"] = field(default_factory=list)
    parent: "StructureNode | None" = None

    def add_child(self, child: "StructureNode") -> None:
        child.parent = self
        self.children.append(child)


class StructureModel:
    def __init__(self) -> None:
        self.root = StructureNode(level="", description="ROOT", part_number="")

    @staticmethod
    def parse_level_code(level: str) -> tuple:
        level_text = str(level).strip()
        if not level_text:
            return tuple()

        result = []
        for token in level_text.split("."):
            token = token.strip()
            if not token:
                continue
            if token.isdigit():
                result.append(int(token))
            else:
                result.append(token)
        return tuple(result)

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> "StructureModel":
        required = ["Level", "Description", "Part Number"]
        for col in required:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")

        model = cls()
        by_code = {tuple(): model.root}

        rows = []
        for _, row in df.iterrows():
            level = "" if pd.isna(row["Level"]) else str(row["Level"]).strip()
            if not level:
                continue
            rows.append(
                {
                    "level": level,
                    "code": cls.parse_level_code(level),
                    "description": "" if pd.isna(row["Description"]) else str(row["Description"]).strip(),
                    "part": "" if pd.isna(row["Part Number"]) else str(row["Part Number"]).strip(),
                }
            )

        rows.sort(key=lambda r: r["code"])

        for row in rows:
            code = row["code"]
            search = code[:-1]
            parent = None
            while search:
                if search in by_code:
                    parent = by_code[search]
                    break
                search = search[:-1]
            if parent is None:
                parent = model.root

            node = StructureNode(
                level=row["level"],
                description=row["description"],
                part_number=row["part"],
            )
            parent.add_child(node)
            by_code[code] = node

        if not model.root.children:
            raise ValueError("No valid rows found in the structure file.")

        return model

    def to_dataframe(self) -> pd.DataFrame:
        rows = []

        def walk(nodes: list[StructureNode], prefix: list[int]) -> None:
            for idx, node in enumerate(nodes, start=1):
                current = prefix + [idx]
                level = ".".join(str(x) for x in current)
                rows.append(
                    {
                        "Level": level,
                        "Description": node.description,
                        "Part Number": node.part_number,
                    }
                )
                walk(node.children, current)

        walk(self.root.children, [])
        return pd.DataFrame(rows)


class ReorderWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"Structure Reorder Helper v{APP_VERSION}")
        self.root.geometry("1000x650")

        self.model: StructureModel | None = None
        self.source_file: str | None = None
        self.item_lookup: dict[str, StructureNode] = {}
        self.undo_stack: list[callable] = []

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        tools = ttk.Frame(main)
        tools.pack(fill="x", pady=10)

        ttk.Button(tools, text="Open Structure", command=self.load_structure).pack(side="left")
        ttk.Button(tools, text="Save As", command=self.save_structure).pack(side="left", padx=8)

        info = (
            "Standalone structure editor for existing Excel files "
            "(Level/Description/Part Number). Reorder levels, add/edit/remove items, "
            "then save a new structure file."
        )
        ttk.Label(main, text=info, justify="left").pack(anchor="w", pady=8)

        tree_frame = ttk.Frame(main)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("part",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="Description")
        self.tree.heading("part", text="Part Number")
        self.tree.column("#0", width=700, anchor="w")
        self.tree.column("part", width=220, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=10)

        self.add_top_btn = ttk.Button(actions, text="Add Top Level", command=self.on_add_top_level)
        self.add_child_btn = ttk.Button(actions, text="Add Child", command=self.on_add_child)
        self.add_sibling_btn = ttk.Button(actions, text="Add Sibling", command=self.on_add_sibling)
        self.edit_btn = ttk.Button(actions, text="Edit Item", command=self.on_edit_item)
        self.move_up_btn = ttk.Button(actions, text="Move Up", command=self.on_move_up)
        self.move_down_btn = ttk.Button(actions, text="Move Down", command=self.on_move_down)
        self.remove_btn = ttk.Button(actions, text="Remove Level", command=self.on_remove)
        self.undo_btn = ttk.Button(actions, text="Undo Remove", command=self.on_undo)
        self.expand_btn = ttk.Button(actions, text="Expand All", command=self.expand_all)
        self.collapse_btn = ttk.Button(actions, text="Collapse All", command=self.collapse_all)

        self.add_top_btn.pack(side="left")
        self.add_child_btn.pack(side="left")
        self.add_sibling_btn.pack(side="left", padx=8)
        self.edit_btn.pack(side="left", padx=8)
        self.move_up_btn.pack(side="left")
        self.move_down_btn.pack(side="left", padx=8)
        self.remove_btn.pack(side="left", padx=8)
        self.undo_btn.pack(side="left", padx=8)
        self.expand_btn.pack(side="left", padx=20)
        self.collapse_btn.pack(side="left", padx=8)

        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.update_buttons())
        self.root.bind("<Delete>", lambda _e: self.on_remove())
        self.root.bind("<Control-z>", lambda _e: self.on_undo())
        self.update_buttons()

    def load_structure(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select structure Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.model = StructureModel.from_dataframe(df)
            self.source_file = file_path
            self.undo_stack.clear()
            self.populate_tree()
            self.apply_default_tree_layout()
            self.update_buttons()
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open structure file:\n{exc}")

    def save_structure(self) -> None:
        if not self.model:
            messagebox.showwarning("No data", "Open a structure file first.")
            return

        initial_name = "reordered_structure.xlsx"
        if self.source_file:
            stem = os.path.splitext(os.path.basename(self.source_file))[0]
            initial_name = f"{stem}_reordered.xlsx"

        output_path = filedialog.asksaveasfilename(
            title="Save reordered structure",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return

        try:
            self.model.to_dataframe().to_excel(output_path, index=False)
            messagebox.showinfo("Saved", f"Saved reordered structure to:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Error", f"Could not save file:\n{exc}")

    def populate_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.item_lookup.clear()

        if not self.model:
            return

        def add_nodes(parent_id: str, nodes: list[StructureNode]) -> None:
            for node in nodes:
                item_id = self.tree.insert(parent_id, "end", text=node.description, values=(node.part_number,), open=True)
                self.item_lookup[item_id] = node
                add_nodes(item_id, node.children)

        add_nodes("", self.model.root.children)

    def selected_node(self) -> StructureNode | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return self.item_lookup.get(sel[0])

    def update_buttons(self) -> None:
        node = self.selected_node()
        can_edit = node is not None and node.parent is not None

        self.add_top_btn.config(state="normal" if self.model else "disabled")
        self.add_child_btn.config(state="normal" if self.model and node else "disabled")
        self.add_sibling_btn.config(state="normal" if can_edit else "disabled")
        self.edit_btn.config(state="normal" if can_edit else "disabled")

        if not can_edit:
            self.move_up_btn.config(state="disabled")
            self.move_down_btn.config(state="disabled")
            self.remove_btn.config(state="disabled")
        else:
            siblings = node.parent.children
            idx = siblings.index(node)
            self.move_up_btn.config(state="normal" if idx > 0 else "disabled")
            self.move_down_btn.config(state="normal" if idx < len(siblings) - 1 else "disabled")
            self.remove_btn.config(state="normal")

        self.undo_btn.config(state="normal" if self.undo_stack else "disabled")

    def prompt_item_values(
        self,
        title: str,
        initial_description: str = "",
        initial_part_number: str = "",
    ) -> tuple[str, str] | None:
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Description").grid(row=0, column=0, sticky="w", pady=(0, 4))
        description_var = tk.StringVar(value=initial_description)
        description_entry = ttk.Entry(frame, textvariable=description_var, width=60)
        description_entry.grid(row=1, column=0, sticky="ew", pady=8)

        ttk.Label(frame, text="Part Number").grid(row=2, column=0, sticky="w", pady=(0, 4))
        part_var = tk.StringVar(value=initial_part_number)
        part_entry = ttk.Entry(frame, textvariable=part_var, width=60)
        part_entry.grid(row=3, column=0, sticky="ew")

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, sticky="e", pady=(12, 0))

        result: tuple[str, str] | None = None

        def on_ok() -> None:
            nonlocal result
            description = description_var.get().strip()
            if not description:
                messagebox.showwarning("Missing Description", "Description is required.", parent=dialog)
                return
            result = (description, part_var.get().strip())
            dialog.destroy()

        def on_cancel() -> None:
            dialog.destroy()

        ttk.Button(buttons, text="Cancel", command=on_cancel).pack(side="right")
        ttk.Button(buttons, text="OK", command=on_ok).pack(side="right", padx=8)

        description_entry.focus_set()
        description_entry.selection_range(0, "end")
        dialog.bind("<Return>", lambda _e: on_ok())
        dialog.bind("<Escape>", lambda _e: on_cancel())
        self.root.wait_window(dialog)
        return result

    def refresh_and_select(self, node: StructureNode | None = None) -> None:
        self.populate_tree()
        self.expand_all()

        if node is not None:
            for item_id, found in self.item_lookup.items():
                if found is node:
                    self.tree.selection_set(item_id)
                    self.tree.focus(item_id)
                    self.tree.see(item_id)
                    break

        self.update_buttons()

    def apply_default_tree_layout(self) -> None:
        for top_item in self.tree.get_children(""):
            self.tree.item(top_item, open=True)
            for child in self.tree.get_children(top_item):
                self.tree.item(child, open=False)

    def on_add_top_level(self) -> None:
        if not self.model:
            return

        values = self.prompt_item_values("Add Top Level")
        if values is None:
            return

        description, part_number = values
        node = StructureNode(level="", description=description, part_number=part_number)
        self.model.root.add_child(node)
        self.refresh_and_select(node)

    def on_add_child(self) -> None:
        if not self.model:
            return

        parent = self.selected_node()
        if parent is None:
            messagebox.showwarning("Select Item", "Select a parent level to add a child.")
            return

        values = self.prompt_item_values("Add Child")
        if values is None:
            return

        description, part_number = values
        node = StructureNode(level="", description=description, part_number=part_number)
        parent.add_child(node)
        self.refresh_and_select(node)

    def on_add_sibling(self) -> None:
        node = self.selected_node()
        if not node or not node.parent:
            return

        values = self.prompt_item_values("Add Sibling")
        if values is None:
            return

        description, part_number = values
        sibling = StructureNode(level="", description=description, part_number=part_number, parent=node.parent)
        siblings = node.parent.children
        siblings.insert(siblings.index(node) + 1, sibling)
        self.refresh_and_select(sibling)

    def on_edit_item(self) -> None:
        node = self.selected_node()
        if not node or not node.parent:
            return

        values = self.prompt_item_values(
            "Edit Item",
            initial_description=node.description,
            initial_part_number=node.part_number,
        )
        if values is None:
            return

        node.description, node.part_number = values
        self.refresh_and_select(node)

    def on_move_up(self) -> None:
        node = self.selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        if idx == 0:
            return
        siblings[idx - 1], siblings[idx] = siblings[idx], siblings[idx - 1]
        self.refresh_and_select(node)

    def on_move_down(self) -> None:
        node = self.selected_node()
        if not node or not node.parent:
            return
        siblings = node.parent.children
        idx = siblings.index(node)
        if idx >= len(siblings) - 1:
            return
        siblings[idx + 1], siblings[idx] = siblings[idx], siblings[idx + 1]
        self.refresh_and_select(node)

    def on_remove(self) -> None:
        node = self.selected_node()
        if not node or not node.parent:
            return

        answer = messagebox.askyesno("Remove Level", f"Remove '{node.description}' and all children?")
        if not answer:
            return

        parent = node.parent
        idx = parent.children.index(node)
        parent.children.pop(idx)

        def undo() -> None:
            parent.children.insert(idx, node)
            self.refresh_and_select(node)

        self.undo_stack.append(undo)
        self.refresh_and_select(parent)

    def on_undo(self) -> None:
        if not self.undo_stack:
            return
        undo_fn = self.undo_stack.pop()
        undo_fn()
        self.update_buttons()

    def expand_all(self) -> None:
        def recurse(item_id: str) -> None:
            self.tree.item(item_id, open=True)
            for child in self.tree.get_children(item_id):
                recurse(child)

        for item_id in self.tree.get_children(""):
            recurse(item_id)

    def collapse_all(self) -> None:
        def recurse(item_id: str) -> None:
            self.tree.item(item_id, open=False)
            for child in self.tree.get_children(item_id):
                recurse(child)

        for item_id in self.tree.get_children(""):
            self.tree.item(item_id, open=True)
            for child in self.tree.get_children(item_id):
                recurse(child)


def main() -> None:
    root = tk.Tk()
    app = ReorderWindow(root)
    app.update_buttons()
    root.mainloop()


if __name__ == "__main__":
    main()
