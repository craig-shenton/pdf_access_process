from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Iterable, List, Optional

import pandas as pd

from output.review_writer import (
    REVIEW_COMMENT_COLUMN,
    REVIEW_STATUS_COLUMN,
    APPROVED_VALUE,
    PENDING_VALUE,
)


DEFAULT_STATUSES = [PENDING_VALUE, APPROVED_VALUE, "REJECTED"]


class ReviewPanel(ttk.Frame):
    """Display and edit review data in a Treeview grid."""

    def __init__(self, master: tk.Widget, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self.columns: List[str] = []
        self.current_df: Optional[pd.DataFrame] = None
        self._save_callback: Optional[Callable[[], None]] = None
        self._edit_widget: Optional[tk.Widget] = None
        self._edit_var: Optional[tk.StringVar] = None
        self._dirty: bool = False

        self._build_widgets()

    def _build_widgets(self) -> None:
        tree_container = ttk.Frame(self)
        tree_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_container,
            show="headings",
            selectmode="extended",
        )
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._start_edit)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        controls = ttk.LabelFrame(self, text="Review controls")
        controls.pack(fill="x", pady=(8, 0))

        status_frame = ttk.Frame(controls)
        status_frame.pack(fill="x", padx=6, pady=(6, 0))

        ttk.Label(status_frame, text="Status:").pack(side="left")
        self.status_var = tk.StringVar(value=PENDING_VALUE)
        self.status_combo = ttk.Combobox(
            status_frame,
            textvariable=self.status_var,
            state="readonly",
            values=DEFAULT_STATUSES,
            width=16,
        )
        self.status_combo.pack(side="left", padx=(4, 12))

        ttk.Label(status_frame, text="Comment:").pack(side="left")
        self.comment_text = tk.Text(status_frame, height=3, width=40, wrap="word")
        self.comment_text.pack(side="left", fill="x", expand=True)

        action_frame = ttk.Frame(controls)
        action_frame.pack(fill="x", padx=6, pady=6)

        self.apply_btn = ttk.Button(
            action_frame,
            text="Apply to selected rows",
            command=self.apply_review_updates,
        )
        self.apply_btn.pack(side="left")

        self.save_btn = ttk.Button(
            action_frame,
            text="Save review CSV",
            command=self._on_save_clicked,
        )
        self.save_btn.pack(side="right")

    def set_save_callback(self, callback: Callable[[], None]) -> None:
        """Register a callback executed when the save button is pressed."""

        self._save_callback = callback

    # Public API ---------------------------------------------------------
    def load_dataframe(self, df: pd.DataFrame) -> None:
        self._teardown_editor()
        self.current_df = df.reset_index(drop=True).copy()
        self.columns = list(self.current_df.columns)
        self._refresh_tree()
        self._update_status_options()
        self._clear_controls()
        self._dirty = False

    def clear(self) -> None:
        self._teardown_editor()
        self.columns = []
        self.current_df = None
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree.configure(columns=self.columns)
        self._update_status_options()
        self._clear_controls()
        self._dirty = False

    def has_data(self) -> bool:
        return self.current_df is not None

    def get_dataframe(self) -> pd.DataFrame:
        if self.current_df is None:
            raise ValueError("No review data loaded")
        # Return a copy to avoid accidental external mutation
        return self.current_df.copy()

    def is_dirty(self) -> bool:
        return self._dirty

    def mark_clean(self) -> None:
        self._dirty = False

    # UI interaction helpers --------------------------------------------
    def apply_review_updates(self) -> None:
        if self.current_df is None:
            return
        selected = self.tree.selection()
        if not selected:
            return

        status_value = self.status_var.get()
        comment_value = self.comment_text.get("1.0", "end-1c")

        for iid in selected:
            index = int(iid)
            if REVIEW_STATUS_COLUMN in self.current_df.columns:
                self.current_df.at[index, REVIEW_STATUS_COLUMN] = status_value
                self.tree.set(iid, REVIEW_STATUS_COLUMN, status_value)
            if REVIEW_COMMENT_COLUMN in self.current_df.columns:
                self.current_df.at[index, REVIEW_COMMENT_COLUMN] = comment_value
                self.tree.set(iid, REVIEW_COMMENT_COLUMN, comment_value)

        if selected:
            self._dirty = True

    # Internal methods --------------------------------------------------
    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree.configure(columns=self.columns)
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w", width=140, stretch=True)

        if self.current_df is None or self.current_df.empty:
            return

        for idx, row in self.current_df.iterrows():
            values = [row.get(col, "") for col in self.columns]
            self.tree.insert("", "end", iid=str(idx), values=values)

    def _update_status_options(self) -> None:
        statuses: Iterable[str] = []
        if self.current_df is not None and REVIEW_STATUS_COLUMN in self.current_df.columns:
            statuses = [
                str(value)
                for value in self.current_df[REVIEW_STATUS_COLUMN].unique()
                if str(value)
            ]
        options = list(dict.fromkeys(DEFAULT_STATUSES + list(statuses)))
        self.status_combo.configure(values=options)
        if options:
            if self.status_var.get() not in options:
                self.status_var.set(options[0])

    def _clear_controls(self) -> None:
        self.status_var.set(PENDING_VALUE)
        self.comment_text.delete("1.0", "end")

    def _on_select(self, _event: tk.Event) -> None:
        if self.current_df is None:
            return
        selected = self.tree.selection()
        if not selected:
            self._clear_controls()
            return
        iid = selected[0]
        index = int(iid)
        if REVIEW_STATUS_COLUMN in self.current_df.columns:
            status_value = str(self.current_df.at[index, REVIEW_STATUS_COLUMN])
            if status_value and status_value not in self.status_combo["values"]:
                self.status_combo.configure(values=list(self.status_combo["values"]) + [status_value])
            if status_value:
                self.status_var.set(status_value)
        if REVIEW_COMMENT_COLUMN in self.current_df.columns:
            comment_value = str(self.current_df.at[index, REVIEW_COMMENT_COLUMN])
            self.comment_text.delete("1.0", "end")
            if comment_value:
                self.comment_text.insert("1.0", comment_value)

    def _start_edit(self, event: tk.Event) -> None:
        if self.current_df is None:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id or not column_id:
            return

        col_index = int(column_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.columns):
            return
        column_name = self.columns[col_index]
        if column_name in (REVIEW_STATUS_COLUMN, REVIEW_COMMENT_COLUMN):
            # Use dedicated controls for these fields
            return

        bbox = self.tree.bbox(row_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        value = self.tree.set(row_id, column_name)
        self._teardown_editor()
        self._edit_var = tk.StringVar(value=value)
        entry = ttk.Entry(self.tree, textvariable=self._edit_var)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()
        entry.bind("<Return>", lambda e: self._finish_edit(row_id, column_name))
        entry.bind("<FocusOut>", lambda e: self._finish_edit(row_id, column_name))
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        self._edit_widget = entry

    def _finish_edit(self, row_id: str, column_name: str) -> None:
        if self._edit_widget is None or self._edit_var is None or self.current_df is None:
            return
        new_value = self._edit_var.get()
        self.tree.set(row_id, column_name, new_value)
        index = int(row_id)
        self.current_df.at[index, column_name] = new_value
        self._dirty = True
        self._teardown_editor()

    def _cancel_edit(self) -> None:
        self._teardown_editor()

    def _teardown_editor(self) -> None:
        if self._edit_widget is not None:
            self._edit_widget.destroy()
        self._edit_widget = None
        self._edit_var = None

    def _on_save_clicked(self) -> None:
        if self._save_callback:
            self._save_callback()
