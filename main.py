"""
SourceThatDeal — PE Deal Sourcing Automator
Entry point and main layout: sidebar navigation + content area.
"""

from nicegui import ui

import asyncio
from datetime import datetime

from rich_editor import RichEditor
from file_manager import (
    ROOT_FOLDER,
    create_folder,
    delete_folder,
    delete_template,
    list_folders,
    move_template,
    list_templates,
    read_template,
    write_template,
    delete_contact_list,
    import_contact_list,
    list_contact_lists,
    read_contact_list,
    write_contact_list,
    list_campaigns,
    read_campaign,
    write_campaign,
    delete_campaign,
    _sanitize_campaign_filename,
)
from email_engine import OutlookClient, _USE_REAL_OUTLOOK
from campaign_engine import (
    extract_template_fields,
    merge_template,
    merge_subject,
    guess_email_column,
    guess_column_match,
    create_campaign_drafts,
)

if _USE_REAL_OUTLOOK:
    try:
        import pythoncom
    except ImportError:
        pass # pythoncom should be available if win32com is

# Corporate palette: White / Slate / Navy (project_plan.md)
NAVY = "#1e3a5f"
SLATE_BG = "#f1f5f9"
SIDEBAR_WIDTH = "16rem"

# Merge-field placeholders for templates (label → tag inserted into HTML)
PLACEHOLDERS = {
    "First Name": "{{First Name}}",
    "Last Name": "{{Last Name}}",
    "Company": "{{Company}}",
    "Title": "{{Title}}",
    "Email": "{{Email}}",
    "Wildcard": None,  # prompts user for a custom field name
}

# Font and size options for editor formatting dropdowns
FONT_OPTIONS = {
    "Arial": "Arial",
    "Verdana": "Verdana",
    "Georgia": "Georgia",
    "Times New Roman": "Times New Roman",
    "Courier New": "Courier New",
}
SIZE_OPTIONS = {"1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7}


def _nav_button(label: str, icon: str, is_active: bool, on_click):
    """Single sidebar nav item with icon; highlights when active."""
    btn = ui.button(on_click=on_click).props("flat no-caps align=left").classes(
        "w-full justify-start text-left rounded-lg py-3 px-4 text-white"
    )
    if is_active:
        btn.classes("bg-blue-600")
    else:
        btn.classes("bg-transparent hover:bg-white/10")
    with btn:
        ui.icon(icon).classes("mr-3")
        ui.label(label)
    return btn


def _build_sidebar(on_nav, active_section: str):
    """Build the left sidebar: branding + Campaigns, Contact Lists, Template Library."""
    with ui.column().classes(
        f"w-[{SIDEBAR_WIDTH}] flex-shrink-0 h-full overflow-hidden"
    ).style(f"background-color: {NAVY};"):
        ui.label("SourceThatDeal").classes(
            "text-white text-xl font-semibold px-4 pt-6 pb-8"
        )
        _nav_button(
            "Campaigns",
            "campaign",
            active_section == "campaigns",
            lambda: on_nav("campaigns"),
        ).classes("mx-2")
        _nav_button(
            "Contact Lists",
            "people",
            active_section == "contact_lists",
            lambda: on_nav("contact_lists"),
        ).classes("mx-2")
        _nav_button(
            "Template Library",
            "library_books",
            active_section == "template_library",
            lambda: on_nav("template_library"),
        ).classes("mx-2")


def _placeholder_content(title: str, subtitle: str):
    """Minimal placeholder for a section (you will replace with real feature UI)."""
    ui.label(title).classes("text-2xl font-semibold text-slate-800 mb-2")
    ui.label(subtitle).classes("text-slate-500")


def _render_template_library(state: dict, refresh):
    """
    Build Template Library: left = folders + template list, right = action buttons + WYSIWYG editor.
    refresh() re-renders this view (e.g. after selection or delete).
    """
    selected = state.get("selected_template")
    folders = list_folders()

    with ui.row().classes("w-full flex-1 min-h-0 min-w-0 gap-0"):
        # Left column: Templates heading, Add new template, New folder, then folder list + templates
        with ui.column().classes(
            "w-60 flex-shrink-0 overflow-auto border-r border-slate-200 pr-4"
        ).style(f"background-color: {SLATE_BG};"):
            ui.label("Templates").classes(
                "text-lg font-semibold text-slate-800 mb-3"
            )
            add_btn = ui.button(on_click=lambda: _add_new_template(state, refresh))
            add_btn.classes("w-full bg-slate-200 text-slate-800 hover:bg-slate-300 mb-2")
            add_btn.props("no-caps flat")
            with add_btn:
                ui.icon("add").classes("mr-2")
                ui.label("Add new template")
            new_folder_btn = ui.button(on_click=lambda: _add_new_folder(state, refresh))
            new_folder_btn.classes("w-full bg-slate-100 text-slate-700 hover:bg-slate-200 mb-3")
            new_folder_btn.props("no-caps flat")
            with new_folder_btn:
                ui.icon("create_new_folder").classes("mr-2")
                ui.label("New folder")
            if not folders:
                ui.label("No templates yet").classes("text-slate-500 text-sm")
            else:
                for folder in folders:
                    with ui.row().classes("w-full items-center mt-2 mb-0.5 gap-0"):
                        ui.label(folder).classes(
                            "text-sm font-medium text-slate-600 flex-1"
                        )
                        if folder != ROOT_FOLDER:
                            del_folder_btn = ui.button(
                                on_click=lambda f=folder: _confirm_delete_folder(f, state, refresh),
                            )
                            del_folder_btn.props("flat round dense size=xs")
                            del_folder_btn.classes("text-slate-400 hover:text-red-500")
                            with del_folder_btn:
                                ui.icon("close").classes("text-xs")
                    for name in list_templates(folder):
                        template_id = f"{folder}/{name}"
                        is_selected = template_id == selected
                        btn = ui.button(
                            on_click=lambda tid=template_id: _select_template(state, tid, refresh)
                        )
                        btn.classes(
                            "w-full justify-start text-left rounded py-1 px-3 text-slate-700"
                        )
                        if is_selected:
                            btn.classes("bg-blue-100 border border-blue-300")
                        else:
                            btn.classes("bg-white border border-slate-200 hover:bg-slate-50")
                        btn.props("flat no-caps align=left dense")
                        with btn:
                            ui.label(name).classes("text-sm")

        # Right column: buttons + editor
        with ui.column().classes("flex-1 min-w-0 flex flex-col pl-4"):
            # Action buttons (top right)
            with ui.row().classes("w-full justify-end gap-2 mb-4"):
                send_btn = ui.button("Send mail", on_click=lambda: _send_mail(state))
                send_btn.classes("bg-blue-600 text-white")
                send_btn.props("no-caps")

                save_btn = ui.button("Save edits", on_click=lambda: _save_template(state))
                save_btn.classes(
                    "bg-blue-50 text-blue-700 border-2 border-blue-600 hover:bg-blue-100"
                )
                save_btn.props("no-caps")
                if not selected:
                    save_btn.props("disable")

                move_btn = ui.button("Move", on_click=lambda: _move_template_dialog(state, refresh))
                move_btn.classes("bg-slate-50 text-slate-700 border border-slate-300")
                move_btn.props("no-caps")
                if not selected:
                    move_btn.props("disable")

                def do_delete():
                    _delete_template(state, refresh)

                del_btn = ui.button("Delete", on_click=do_delete)
                del_btn.classes("bg-red-600 text-white border border-red-700")
                del_btn.props("no-caps")
                if not selected:
                    del_btn.props("disable")

            # Rich WYSIWYG editor (font size, styling, links, lists, etc.; Outlook-friendly HTML)
            initial_html = read_template(selected) if selected else ""
            editor = RichEditor(value=initial_html).classes("w-full flex-1 min-h-64")
            state["template_editor"] = editor

            # Formatting controls + Insert Field dropdown
            with ui.row().classes("w-full items-center gap-2 mt-2"):
                ui.label("Preview / Edit Template Canvas").classes(
                    "text-slate-500 text-sm flex-1"
                )

                # Font family selector
                font_select = ui.select(
                    options=list(FONT_OPTIONS.keys()),
                    label="Font",
                    value=None,
                ).props("outlined dense options-dense").classes("w-36")

                def on_font_selected(e):
                    if e.value:
                        editor.set_font_name(FONT_OPTIONS[e.value])
                        font_select.set_value(None)

                font_select.on_value_change(on_font_selected)

                # Font size selector
                size_select = ui.select(
                    options=list(SIZE_OPTIONS.keys()),
                    label="Size",
                    value=None,
                ).props("outlined dense options-dense").classes("w-24")

                def on_size_selected(e):
                    if e.value:
                        editor.set_font_size(SIZE_OPTIONS[e.value])
                        size_select.set_value(None)

                size_select.on_value_change(on_size_selected)

                # Insert Field dropdown (merge-field placeholders)
                field_select = ui.select(
                    options=list(PLACEHOLDERS.keys()),
                    label="Insert Field",
                    value=None,
                ).props("outlined dense options-dense").classes("w-44")

                def on_field_selected(e):
                    chosen = e.value
                    if chosen is None:
                        return
                    tag = PLACEHOLDERS.get(chosen)
                    if tag is not None:
                        editor.insert_at_cursor(tag)
                        field_select.set_value(None)
                    else:
                        _wildcard_dialog(editor, field_select)

                field_select.on_value_change(on_field_selected)


def _render_contact_lists(state: dict, refresh):
    """
    Build Contact Lists view: left = list sidebar, right = editable AG Grid table.
    """
    selected = state.get("selected_contact_list")  # filename e.g. "Main Insight.xlsx"
    contact_lists = list_contact_lists()

    with ui.row().classes("w-full flex-1 min-h-0 min-w-0 gap-0"):
        # Left column: heading, import/create buttons, list of contact lists
        with ui.column().classes(
            "w-60 flex-shrink-0 overflow-auto border-r border-slate-200 pr-4"
        ).style(f"background-color: {SLATE_BG};"):
            ui.label("Contact Lists").classes(
                "text-lg font-semibold text-slate-800 mb-3"
            )

            # Import button
            def handle_upload(e):
                for file in e.files if hasattr(e, 'files') else [e]:
                    name = getattr(file, 'name', 'upload.xlsx')
                    content = file.content if hasattr(file, 'content') else file.read()
                    import_contact_list(content, name)
                ui.notify("List imported.")
                refresh()

            upload = ui.upload(
                on_upload=handle_upload,
                auto_upload=True,
                label="Import new list",
            ).props('accept=".xlsx" flat').classes(
                "w-full mb-2"
            )
            upload.classes("bg-blue-600 text-white rounded")

            # Create new empty list button
            create_btn = ui.button(on_click=lambda: _create_new_contact_list(state, refresh))
            create_btn.classes("w-full bg-slate-200 text-slate-800 hover:bg-slate-300 mb-3")
            create_btn.props("no-caps flat")
            with create_btn:
                ui.icon("add").classes("mr-2")
                ui.label("Create new list")

            if not contact_lists:
                ui.label("No contact lists yet").classes("text-slate-500 text-sm")
            else:
                for cl in contact_lists:
                    filename = cl["filename"]
                    is_selected = filename == selected
                    with ui.row().classes("w-full items-center gap-0 mb-1"):
                        btn = ui.button(
                            on_click=lambda f=filename: _select_contact_list(state, f, refresh)
                        )
                        btn.classes(
                            "flex-1 justify-start text-left rounded py-1 px-3 text-slate-700"
                        )
                        if is_selected:
                            btn.classes("bg-blue-100 border border-blue-300")
                        else:
                            btn.classes("bg-white border border-slate-200 hover:bg-slate-50")
                        btn.props("flat no-caps align=left dense")
                        with btn:
                            with ui.column().classes("gap-0"):
                                ui.label(cl["name"]).classes("text-sm font-medium")
                                ui.label(f"{cl['row_count']:,} Contacts").classes(
                                    "text-xs text-slate-500"
                                )
                        del_btn = ui.button(
                            on_click=lambda f=filename: _confirm_delete_contact_list(f, state, refresh),
                        )
                        del_btn.props("flat round dense size=xs")
                        del_btn.classes("text-slate-400 hover:text-red-500")
                        with del_btn:
                            ui.icon("delete").classes("text-xs")

        # Right column: action buttons + editable table
        with ui.column().classes("flex-1 min-w-0 flex flex-col pl-4"):
            if not selected:
                ui.label("Select a contact list to view and edit.").classes(
                    "text-slate-500 mt-4"
                )
            else:
                columns, rows = read_contact_list(selected)
                if not columns:
                    columns = ["Email", "First Name", "Last Name", "Company", "Title"]

                # Store rows in state for editing
                state["cl_columns"] = list(columns)
                # Ensure there's always an empty row at the bottom
                indexed_rows = [dict(row, _idx=i) for i, row in enumerate(rows)]
                next_idx = len(indexed_rows)
                empty_row = {col: "" for col in columns}
                empty_row["_idx"] = next_idx
                indexed_rows.append(empty_row)
                state["cl_rows"] = indexed_rows

                # Action buttons
                with ui.row().classes("w-full justify-end gap-2 mb-4"):
                    del_row_btn = ui.button(
                        "Delete selected rows",
                        on_click=lambda: _delete_contact_row(state, table, refresh),
                    )
                    del_row_btn.classes("bg-red-50 text-red-700 border border-red-300")
                    del_row_btn.props("no-caps")

                    save_btn = ui.button(
                        "Save changes",
                        on_click=lambda: _save_contact_list(state, table, refresh),
                    )
                    save_btn.classes(
                        "bg-blue-50 text-blue-700 border-2 border-blue-600 hover:bg-blue-100"
                    )
                    save_btn.props("no-caps")

                # Editable column headers
                ui.label("Column Headers").classes("text-xs text-slate-500 mt-1 mb-1")
                with ui.row().classes("w-full gap-2 mb-2"):
                    col_inputs = []
                    for i, col in enumerate(columns):
                        inp = ui.input(value=col).props("dense outlined").classes("flex-1")
                        col_inputs.append(inp)
                    state["cl_col_inputs"] = col_inputs

                # Quasar QTable
                table_columns = [
                    {"name": col, "label": col, "field": col, "align": "left", "sortable": True}
                    for col in columns
                ]
                table = ui.table(
                    columns=table_columns,
                    rows=state["cl_rows"],
                    row_key="_idx",
                    selection="multiple",
                    pagination={"rowsPerPage": 20},
                ).classes("w-full")
                table.props("flat bordered")

                # Make each cell editable — emit changes back to Python
                for col in columns:
                    safe_col = col.replace("'", "\\'")
                    table.add_slot(
                        f"body-cell-{col}",
                        f'''
                        <q-td :props="props">
                            <q-input v-model="props.row['{safe_col}']" dense borderless
                                input-class="text-sm"
                                @update:model-value="(val) => $parent.$emit('cell_edit', {{idx: props.row._idx, col: '{safe_col}', value: val}})" />
                        </q-td>
                        '''
                    )

                def handle_cell_edit(e):
                    idx = e.args.get("idx")
                    col_name = e.args.get("col")
                    value = e.args.get("value", "")
                    for row in state["cl_rows"]:
                        if row.get("_idx") == idx:
                            row[col_name] = value
                            break
                    # Auto-add empty row if the last row was edited
                    last_row = state["cl_rows"][-1] if state["cl_rows"] else None
                    if last_row and any(last_row.get(c, "") for c in state["cl_columns"]):
                        new_idx = max(r.get("_idx", 0) for r in state["cl_rows"]) + 1
                        new_empty = {c: "" for c in state["cl_columns"]}
                        new_empty["_idx"] = new_idx
                        state["cl_rows"].append(new_empty)
                        table.rows = state["cl_rows"]
                        table.update()

                table.on("cell_edit", handle_cell_edit)

                ui.label("Edit cells directly in the table. Click 'Save changes' to persist.").classes(
                    "text-slate-500 text-sm mt-2"
                )


def _select_contact_list(state: dict, filename: str, refresh):
    state["selected_contact_list"] = filename
    refresh()


def _create_new_contact_list(state: dict, refresh):
    """Dialog to create a new empty contact list."""
    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label("New contact list")
        name_input = ui.input("List name").classes("w-full")
        name_input.props("outlined")
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def create():
                name = (name_input.value or "").strip()
                if not name:
                    ui.notify("Enter a list name.", type="warning") # Moved before dlg.close()
                    return
                default_cols = ["Email", "First Name", "Last Name", "Company", "Title"]
                filename = f"{name}.xlsx"
                write_contact_list(filename, default_cols, [])
                state["selected_contact_list"] = filename
                dlg.close()
                refresh()
                ui.notify("Contact list created.") # Moved before dlg.close()

            ui.button("Create", on_click=create).classes("bg-blue-600 text-white").props("no-caps")
    dlg.open()


def _confirm_delete_contact_list(filename: str, state: dict, refresh):
    """Confirm and delete a contact list."""
    display_name = filename.replace(".xlsx", "")
    with ui.dialog() as dlg, ui.card().classes("p-4"):
        ui.label(f'Delete contact list "{display_name}"? This cannot be undone.')
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def confirm():
                ui.notify("Contact list deleted.") # Moved before dlg.close()
                delete_contact_list(filename)
                if state.get("selected_contact_list") == filename:
                    state["selected_contact_list"] = None
                dlg.close()
                refresh()
            ui.button("Delete", on_click=confirm).classes(
                "bg-red-600 text-white"
            ).props("no-caps")
    dlg.open()


def _delete_contact_row(state: dict, table, refresh):
    """Delete the currently selected row(s) from the table."""
    selected = table.selected
    if not selected:
        ui.notify("Select one or more rows first.", type="warning")
        return
    # Collect all selected indices
    sel_indices = {row.get("_idx") for row in selected}
    rows = state.get("cl_rows", [])
    # Filter out selected rows
    state["cl_rows"] = [r for r in rows if r.get("_idx") not in sel_indices]
    table.rows = state["cl_rows"]
    table.selected = []
    table.update()
    count = len(sel_indices)
    ui.notify(f"{count} row{'s' if count > 1 else ''} deleted.")


def _save_contact_list(state: dict, table, refresh):
    """Save the current table data back to the Excel file."""
    selected = state.get("selected_contact_list")
    if not selected:
        return
    # Check if column headers were renamed
    col_inputs = state.get("cl_col_inputs", [])
    old_columns = state.get("cl_columns", [])
    new_columns = [inp.value.strip() or old_columns[i] for i, inp in enumerate(col_inputs)]
    rows = state.get("cl_rows", [])
    # Remap row keys if columns were renamed
    clean_rows = []
    for r in rows:
        new_row = {}
        is_empty = True
        for old_col, new_col in zip(old_columns, new_columns):
            val = r.get(old_col, "")
            new_row[new_col] = val
            if val:
                is_empty = False
        if not is_empty:
            clean_rows.append(new_row)
    try:
        write_contact_list(selected, new_columns, clean_rows)
        state["cl_columns"] = new_columns
        ui.notify("Contact list saved.")
        refresh()
    except Exception as e:
        ui.notify(f"Failed to save: {e}", type="negative")


# ---------------------------------------------------------------------------
# Campaigns
# ---------------------------------------------------------------------------

STATUS_BADGES = {
    "draft": ("Draft", "bg-slate-200 text-slate-600"),
    "test_sent": ("Tested", "bg-amber-100 text-amber-700"),
    "completed": ("Sent", "bg-green-100 text-green-700"),
}


def _render_campaigns(state: dict, refresh):
    """Build Campaigns view: left = campaign list, right = stepper wizard."""
    selected_filename = state.get("selected_campaign")  # e.g. "Q1_Outreach.json"
    campaigns = list_campaigns()

    with ui.row().classes("w-full flex-1 min-h-0 min-w-0 gap-0"):
        # ---- Left sidebar ----
        with ui.column().classes(
            "w-60 flex-shrink-0 overflow-auto border-r border-slate-200 pr-4"
        ).style(f"background-color: {SLATE_BG};"):
            ui.label("Campaigns").classes("text-lg font-semibold text-slate-800 mb-3")

            create_btn = ui.button(
                on_click=lambda: _create_new_campaign(state, refresh),
            )
            create_btn.classes("w-full bg-blue-600 text-white hover:bg-blue-700 mb-3")
            create_btn.props("no-caps")
            with create_btn:
                ui.icon("add").classes("mr-2")
                ui.label("New campaign")

            if not campaigns:
                ui.label("No campaigns yet").classes("text-slate-500 text-sm")
            else:
                for camp in campaigns:
                    fname = camp["filename"]
                    is_selected = fname == selected_filename
                    with ui.row().classes("w-full items-center gap-0 mb-1"):
                        btn = ui.button(
                            on_click=lambda f=fname: _select_campaign(state, f, refresh),
                        )
                        btn.classes(
                            "flex-1 justify-start text-left rounded py-1 px-3 text-slate-700"
                        )
                        if is_selected:
                            btn.classes("bg-blue-100 border border-blue-300")
                        else:
                            btn.classes("bg-white border border-slate-200 hover:bg-slate-50")
                        btn.props("flat no-caps align=left dense")
                        with btn:
                            with ui.column().classes("gap-0"):
                                ui.label(camp["name"]).classes("text-sm font-medium")
                                badge_label, badge_cls = STATUS_BADGES.get(
                                    camp["status"], STATUS_BADGES["draft"]
                                )
                                ui.badge(badge_label).classes(f"text-xs {badge_cls}")
                        del_btn = ui.button(
                            on_click=lambda f=fname: _confirm_delete_campaign(f, state, refresh),
                        )
                        del_btn.props("flat round dense size=xs")
                        del_btn.classes("text-slate-400 hover:text-red-500")
                        with del_btn:
                            ui.icon("delete").classes("text-xs")

        # ---- Right panel: wizard ----
        with ui.column().classes("flex-1 min-w-0 flex flex-col pl-4"):
            if not selected_filename:
                ui.label("Select or create a campaign to get started.").classes(
                    "text-slate-500 mt-4"
                )
            else:
                campaign = read_campaign(selected_filename)
                if not campaign:
                    ui.label("Campaign not found.").classes("text-red-500 mt-4")
                else:
                    _render_campaign_wizard(state, campaign, refresh)


def _select_campaign(state: dict, filename: str, refresh):
    state["selected_campaign"] = filename
    refresh()


def _create_new_campaign(state: dict, refresh):
    """Dialog to create a new campaign."""
    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label("New campaign").classes("text-lg font-semibold")
        name_input = ui.input("Campaign name").classes("w-full")
        name_input.props("outlined")
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def create():
                name = (name_input.value or "").strip()
                if not name:
                    ui.notify("Enter a campaign name.", type="warning") # Moved before dlg.close()
                    return
                try:
                    filename = _sanitize_campaign_filename(name)
                    data = {
                        "name": name,
                        "template_id": "", # Fix: This was missing a comma
                        "contact_list": "",
                        "email_column": "",
                        "subject_template": "",
                        "field_mapping": {},
                        "deferred_delivery": None,
                        "status": "draft",
                        "drafts_created": 0,
                        "created_at": datetime.now().isoformat(),
                    }
                    write_campaign(filename, data)
                    state["selected_campaign"] = filename
                    dlg.close()
                    refresh()
                    ui.notify("Campaign created.") # Moved before dlg.close()
                except Exception as e:
                    ui.notify(f"Failed: {e}", type="negative")

            ui.button("Create", on_click=create).classes(
                "bg-blue-600 text-white"
            ).props("no-caps")
    dlg.open()


def _confirm_delete_campaign(filename: str, state: dict, refresh):
    """Confirm and delete a campaign."""
    with ui.dialog() as dlg, ui.card().classes("p-4"):
        ui.label("Delete this campaign? This cannot be undone.")
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def confirm():
                ui.notify("Campaign deleted.") # Moved before dlg.close()
                delete_campaign(filename)
                if state.get("selected_campaign") == filename:
                    state["selected_campaign"] = None
                dlg.close()
                refresh()
            ui.button("Delete", on_click=confirm).classes(
                "bg-red-600 text-white"
            ).props("no-caps")
    dlg.open()


def _render_campaign_wizard(state: dict, campaign: dict, refresh):
    """Render the 5-step campaign wizard using Quasar stepper."""
    templates = list_templates()  # flattened template_ids
    contact_lists_meta = list_contact_lists()
    cl_options = {cl["filename"]: f"{cl['name']} ({cl['row_count']:,} rows)" for cl in contact_lists_meta}

    # Mutable refs dict — holds live references to UI elements across steps
    refs = {
        "template_select": None,
        "cl_select": None,
        "email_select": None,
        "subject_input": None,
        "field_selects": {},
        "enable_schedule": None,
        "date_input": None,
        "time_input": None,
    }

    def _collect_values() -> dict:
        """Read current values from all wizard UI elements into a dict."""
        tid = refs["template_select"].value if refs["template_select"] else ""
        cl_fname = refs["cl_select"].value if refs["cl_select"] else ""
        email_col = refs["email_select"].value if refs["email_select"] else ""
        subj = refs["subject_input"].value if refs["subject_input"] else ""
        mapping = {
            f: sel.value
            for f, sel in refs.get("field_selects", {}).items()
            if sel and sel.value
        }
        deferred_str = None
        if (refs["enable_schedule"] and refs["enable_schedule"].value
                and refs["date_input"] and refs["date_input"].value):
            t = refs["time_input"].value if refs["time_input"] else "09:00"
            deferred_str = f"{refs['date_input'].value}T{t or '09:00'}:00"
        return {
            "template_id": tid or "",
            "contact_list": cl_fname or "",
            "email_column": email_col or "",
            "subject_template": subj or "",
            "field_mapping": mapping,
            "deferred_delivery": deferred_str,
        }

    def _save_step_values():
        """Persist current wizard values into the campaign dict."""
        vals = _collect_values()
        campaign.update(vals)

    with ui.stepper().props("vertical animated").classes("w-full") as stepper:
        # ---- Step 1: Select Template ----
        with ui.step("Template"):
            ui.label("Choose an email template for this campaign").classes(
                "text-slate-600 mb-2"
            )
            template_select = ui.select(
                options=templates,
                value=campaign.get("template_id") or None,
                label="Template",
            ).props("outlined").classes("w-80")
            refs["template_select"] = template_select

            # Live preview of selected template
            preview_container = ui.column().classes("w-full mt-2")

            def update_template_preview():
                preview_container.clear()
                tid = template_select.value
                if tid:
                    html = read_template(tid)
                    if html:
                        with preview_container:
                            ui.label("Preview:").classes("text-xs text-slate-500")
                            ui.html(html).classes(
                                "w-full border border-slate-200 rounded p-3 bg-white "
                                "max-h-40 overflow-auto text-sm"
                            )

            template_select.on_value_change(lambda _: update_template_preview())
            update_template_preview()

            with ui.stepper_navigation():
                def step1_next():
                    _save_step_values()
                    stepper.next()
                ui.button("Next", on_click=step1_next).props("no-caps")

        # ---- Step 2: Select Contact List ----
        with ui.step("Contact List"):
            ui.label("Choose a contact list").classes("text-slate-600 mb-2")
            cl_select = ui.select(
                options=cl_options,
                value=campaign.get("contact_list") or None,
                label="Contact List",
            ).props("outlined").classes("w-80")
            refs["cl_select"] = cl_select

            cl_info = ui.column().classes("mt-2")

            def update_cl_info():
                cl_info.clear()
                fname = cl_select.value
                if fname:
                    try:
                        cols, rows = read_contact_list(fname)
                        with cl_info:
                            ui.label(f"{len(rows)} rows, {len(cols)} columns").classes(
                                "text-sm text-slate-500"
                            )
                            ui.label(f"Columns: {', '.join(cols)}").classes(
                                "text-xs text-slate-400"
                            )
                    except Exception:
                        pass

            cl_select.on_value_change(lambda _: update_cl_info())
            update_cl_info()

            with ui.stepper_navigation():
                def step2_next():
                    _save_step_values()
                    build_mapping_ui()  # rebuild mapping based on template + contact list
                    stepper.next()
                ui.button("Next", on_click=step2_next).props("no-caps")
                ui.button("Back", on_click=stepper.previous).props("flat no-caps")

        # ---- Step 3: Field Mapping ----
        with ui.step("Field Mapping"):
            mapping_container = ui.column().classes("w-full")

            def build_mapping_ui():
                mapping_container.clear()
                tid = refs["template_select"].value if refs["template_select"] else ""
                cl_fname = refs["cl_select"].value if refs["cl_select"] else ""
                if not tid or not cl_fname:
                    with mapping_container:
                        ui.label("Please complete previous steps first.").classes(
                            "text-amber-600"
                        )
                    return
                html = read_template(tid)
                fields = extract_template_fields(html)
                try:
                    columns, _ = read_contact_list(cl_fname)
                except Exception:
                    columns = []
                existing_mapping = campaign.get("field_mapping", {})

                with mapping_container:
                    # Email column
                    ui.label("Email address column").classes(
                        "text-sm font-medium text-slate-700 mt-1"
                    )
                    email_sel = ui.select(
                        options=columns,
                        value=campaign.get("email_column") or guess_email_column(columns),
                        label="Column with email addresses",
                    ).props("outlined dense").classes("w-80 mb-4")
                    refs["email_select"] = email_sel

                    # Subject line
                    ui.label("Subject line").classes(
                        "text-sm font-medium text-slate-700"
                    )
                    subj_inp = ui.input(
                        value=campaign.get("subject_template", ""),
                        placeholder="e.g. Introduction - {{Company}}",
                    ).props("outlined dense").classes("w-full mb-1")
                    refs["subject_input"] = subj_inp
                    ui.label(
                        "Use {{Field Name}} for personalization"
                    ).classes("text-xs text-slate-400 mb-4")

                    # Field mapping rows
                    if fields:
                        ui.label("Map template fields to columns").classes(
                            "text-sm font-medium text-slate-700 mb-2"
                        )
                        refs["field_selects"] = {}
                        for field in fields:
                            with ui.row().classes("items-center gap-3 mb-2"):
                                ui.label(f"{{{{{field}}}}}").classes(
                                    "w-44 text-sm font-mono bg-slate-100 px-2 py-1 rounded"
                                )
                                ui.icon("arrow_forward").classes("text-slate-400")
                                sel = ui.select(
                                    options=columns,
                                    value=existing_mapping.get(
                                        field, guess_column_match(field, columns)
                                    ),
                                    label="Column",
                                ).props("outlined dense").classes("w-60")
                                refs["field_selects"][field] = sel
                    else:
                        ui.label("No {{placeholders}} found in this template.").classes(
                            "text-slate-500 text-sm"
                        )

            build_mapping_ui()

            with ui.stepper_navigation():
                def step3_next():
                    _save_step_values()  # capture mapping selections into campaign
                    stepper.next()

                ui.button("Next", on_click=step3_next).props("no-caps")
                ui.button("Back", on_click=stepper.previous).props("flat no-caps")

        # ---- Step 4: Schedule ----
        with ui.step("Schedule"):
            ui.label("Schedule delivery (optional)").classes("text-slate-600 mb-2")
            ui.label(
                "If enabled, drafts will have a deferred delivery time. "
                "When you send them from Outlook, they stay in Outbox until the scheduled time."
            ).classes("text-xs text-slate-400 mb-4")

            has_existing_deferred = bool(campaign.get("deferred_delivery"))
            enable_schedule = ui.checkbox(
                "Enable scheduled delivery",
                value=has_existing_deferred,
            )
            refs["enable_schedule"] = enable_schedule

            existing_date = ""
            existing_time = "09:00"
            if has_existing_deferred:
                try:
                    dt = datetime.fromisoformat(campaign["deferred_delivery"])
                    existing_date = dt.strftime("%Y-%m-%d")
                    existing_time = dt.strftime("%H:%M")
                except (ValueError, TypeError):
                    pass

            with ui.row().classes("gap-4 mt-2") as schedule_row:
                date_input = ui.input(
                    "Date", value=existing_date,
                ).props("outlined dense type=date").classes("w-48")
                time_input = ui.input(
                    "Time", value=existing_time,
                ).props("outlined dense type=time").classes("w-36")
                refs["date_input"] = date_input
                refs["time_input"] = time_input

            schedule_row.bind_visibility_from(enable_schedule, "value")

            with ui.stepper_navigation():
                def step4_next():
                    _save_step_values()  # capture schedule into campaign
                    build_review()       # rebuild review with latest values
                    stepper.next()
                ui.button("Next", on_click=step4_next).props("no-caps")
                ui.button("Back", on_click=stepper.previous).props("flat no-caps")

        # ---- Step 5: Review & Create ----
        with ui.step("Review & Create"):
            review_container = ui.column().classes("w-full")

            def build_review():
                review_container.clear()
                # Read from campaign dict (updated by _save_step_values at each step)
                tid = campaign.get("template_id", "")
                cl_fname = campaign.get("contact_list", "")
                email_col = campaign.get("email_column", "")
                subj_tmpl = campaign.get("subject_template", "")
                field_mapping = campaign.get("field_mapping", {})
                deferred_str = campaign.get("deferred_delivery")

                with review_container:
                    # Summary
                    with ui.card().classes("w-full p-4 mb-4 bg-white"):
                        ui.label("Campaign Summary").classes(
                            "text-sm font-semibold text-slate-700 mb-2"
                        )
                        ui.label(f"Template: {tid or 'Not selected'}").classes("text-sm")
                        cl_display = cl_options.get(cl_fname, cl_fname or "Not selected")
                        ui.label(f"Contact List: {cl_display}").classes("text-sm")
                        ui.label(f"Email Column: {email_col or 'Not set'}").classes("text-sm")
                        ui.label(f"Subject: {subj_tmpl or '(empty)'}").classes("text-sm")
                        if field_mapping:
                            for f, col in field_mapping.items():
                                ui.label(f"  {{{{{f}}}}} → {col}").classes(
                                    "text-xs text-slate-500 font-mono"
                                )
                        if deferred_str:
                            ui.label(f"Scheduled: {deferred_str}").classes(
                                "text-sm text-blue-600 mt-1"
                            )

                    # Preview first row
                    if tid and cl_fname:
                        try:
                            html = read_template(tid)
                            columns, rows = read_contact_list(cl_fname)
                            if rows:
                                first_row = rows[0]
                                preview_html = merge_template(html, first_row, field_mapping)
                                preview_subj = merge_subject(subj_tmpl, first_row, field_mapping)
                                to_addr = first_row.get(email_col, "")
                                ui.label("Preview (first contact)").classes(
                                    "text-sm font-medium text-slate-700 mb-1"
                                )
                                ui.label(f"To: {to_addr}").classes("text-xs text-slate-500")
                                ui.label(f"Subject: {preview_subj}").classes(
                                    "text-xs text-slate-500 mb-2"
                                )
                                ui.html(preview_html).classes(
                                    "w-full border border-slate-200 rounded p-4 bg-white "
                                    "max-h-60 overflow-auto"
                                )
                        except Exception:
                            pass

                    # Progress elements (hidden until creation starts)
                    progress_label = ui.label("").classes("text-sm text-slate-500 mt-4")
                    progress_bar = ui.linear_progress(value=0, show_value=False).classes(
                        "w-full"
                    )
                    progress_bar.set_visibility(False)

                    # Action buttons — read live from campaign dict
                    with ui.row().classes("gap-4 mt-4"):
                        ui.button(
                            "Create test draft",
                            on_click=lambda: _run_campaign_drafts(
                                state, campaign,
                                progress_label, progress_bar, refresh,
                                test_only=True,
                            ),
                        ).classes(
                            "bg-blue-50 text-blue-700 border-2 border-blue-600"
                        ).props("no-caps")

                        ui.button(
                            "Create all drafts",
                            on_click=lambda: _run_campaign_drafts(
                                state, campaign,
                                progress_label, progress_bar, refresh,
                                test_only=False,
                            ),
                        ).classes("bg-blue-600 text-white").props("no-caps")

                    # Save campaign config
                    with ui.row().classes("mt-4"):
                        ui.button(
                            "Save campaign",
                            on_click=lambda: _save_campaign_from_wizard(
                                state, campaign, refresh,
                            ),
                        ).classes("bg-slate-200 text-slate-700").props("flat no-caps")

            build_review()

            with ui.stepper_navigation():
                ui.button("Back", on_click=stepper.previous).props("flat no-caps")


def _save_campaign_from_wizard(state, campaign, refresh):
    """Persist the campaign dict to JSON."""
    filename = campaign.get("filename") or state.get("selected_campaign")
    if filename:
        write_campaign(filename, campaign)
        ui.notify("Campaign saved.")


async def _run_campaign_drafts(
    state, campaign,
    progress_label, progress_bar, refresh,
    test_only=False,
):
    """Create drafts (test or all) with progress feedback. Reads from campaign dict."""
    template_id = campaign.get("template_id", "")
    contact_list = campaign.get("contact_list", "")
    email_column = campaign.get("email_column", "")
    subject_template = campaign.get("subject_template", "")
    field_mapping = campaign.get("field_mapping", {})
    deferred_str = campaign.get("deferred_delivery")

    if not template_id:
        ui.notify("Select a template first.", type="warning")
        return
    if not contact_list:
        ui.notify("Select a contact list first.", type="warning")
        return
    if not email_column:
        ui.notify("Set the email column first.", type="warning")
        return

    # Save config first
    _save_campaign_from_wizard(state, campaign, refresh)

    outlook = OutlookClient()
    html = read_template(template_id)
    try:
        _, rows = read_contact_list(contact_list)
    except Exception as e:
        ui.notify(f"Cannot read contact list: {e}", type="negative")
        return

    if not rows:
        ui.notify("Contact list is empty.", type="warning")
        return

    if test_only:
        rows = [rows[0]]

    # Parse deferred delivery
    deferred_dt = None
    if deferred_str:
        try:
            deferred_dt = datetime.fromisoformat(deferred_str)
        except (ValueError, TypeError):
            pass

    # Cache Outlook signature once
    signature_html = outlook.get_default_signature_html()

    total = len(rows)
    progress_state = {"current": 0, "done": False, "success": 0, "errors": []}

    progress_bar.set_visibility(True)
    progress_label.set_text(f"Creating drafts... 0/{total}")

    def do_creation():
        if _USE_REAL_OUTLOOK:
            pythoncom.CoInitialize()  # Initialize COM for this thread
            # Instantiate OutlookClient and get signature within the worker thread
            # This ensures COM objects are bound to this thread.
            thread_outlook_client = OutlookClient()
            signature_html = thread_outlook_client.get_default_signature_html()
        else:
            # For mock backend, instantiate here for consistency
            thread_outlook_client = OutlookClient()
            signature_html = ""  # Mock backend doesn't have a real signature
        def on_progress(current, tot):
            progress_state["current"] = current

        success, errors = create_campaign_drafts(
            outlook=thread_outlook_client,  # Pass the thread-local client
            template_html=html,
            subject_template=subject_template or "",
            rows=rows,
            field_mapping=field_mapping or {},
            email_column=email_column,
            deferred_delivery=deferred_dt,
            signature_html=signature_html,
            on_progress=on_progress,
        )
        progress_state["success"] = success
        progress_state["errors"] = errors
        progress_state["done"] = True
        if _USE_REAL_OUTLOOK:
            pythoncom.CoUninitialize() # Uninitialize COM

    # Poll progress with a timer
    def poll_progress():
        cur = progress_state["current"]
        if total > 0:
            progress_bar.set_value(cur / total)
            progress_label.set_text(f"Creating drafts... {cur}/{total}")
        if progress_state["done"]:
            timer.deactivate()
            progress_bar.set_visibility(False)
            s = progress_state["success"]
            errs = progress_state["errors"]
            if test_only:
                if s:
                    ui.notify("Test draft created — check your Outlook Drafts.", type="positive")
                    campaign["status"] = "test_sent"
                else:
                    ui.notify(f"Test failed: {'; '.join(errs)}", type="negative")
            else:
                campaign["status"] = "completed"
                campaign["drafts_created"] = s
                if errs:
                    ui.notify(f"{s} drafts created, {len(errs)} errors.", type="warning")
                else:
                    ui.notify(f"All {s} drafts created!", type="positive")
            # Persist updated status
            fname = campaign.get("filename") or state.get("selected_campaign")
            if fname:
                write_campaign(fname, campaign)
            progress_label.set_text(f"Done: {s}/{total} drafts created.")

    timer = ui.timer(0.3, poll_progress)

    # Run creation in background thread
    loop = asyncio.get_event_loop()
    await loop.run_in_executor(None, do_creation)


def _wildcard_dialog(editor: RichEditor, field_select):
    """Open dialog for custom wildcard field name, then insert {{name}} at cursor."""
    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label("Custom Wildcard")
        name_input = ui.input("Field name (e.g. Fund Size)").classes("w-full")
        name_input.props("outlined")
        with ui.row().classes("mt-4 gap-2"):
            def cancel():
                field_select.set_value(None)
                dlg.close()

            ui.button("Cancel", on_click=cancel).props("flat no-caps")

            def insert():
                ui.notify("Wildcard inserted.") # Moved before dlg.close()
                name = (name_input.value or "").strip()
                if not name:
                    ui.notify("Enter a field name.", type="warning")
                    return
                editor.insert_at_cursor(f"{{{{{name}}}}}")
                field_select.set_value(None)
                dlg.close()
            ui.button("Insert", on_click=insert).classes("bg-blue-600 text-white").props("no-caps")
    dlg.open()


def _select_template(state: dict, name: str, refresh):
    state["selected_template"] = name
    refresh()


def _add_new_folder(state: dict, refresh):
    """Open dialog to enter new folder name; create folder and refresh list."""
    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label("New folder")
        name_input = ui.input("Folder name").classes("w-full")
        name_input.props("outlined")
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def create():
                name = (name_input.value or "").strip()
                if not name:
                    ui.notify("Enter a folder name.", type="warning")
                    return
                try:
                    ui.notify("Folder created.") # Moved before dlg.close()
                    create_folder(name)
                    dlg.close()
                    refresh()
                except Exception as e:
                    ui.notify(str(e), type="negative")
            ui.button("Create", on_click=create).classes("bg-blue-600 text-white").props("no-caps")
    dlg.open()


def _confirm_delete_folder(folder_name: str, state: dict, refresh):
    """Confirm and delete a folder."""
    with ui.dialog() as dlg, ui.card().classes("p-4"):
        ui.label(f'Delete folder "{folder_name}"?')
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def confirm():
                try:
                    ui.notify("Folder deleted.") # Moved before dlg.close()
                    delete_folder(folder_name)
                    if state.get("selected_template", "").startswith(f"{folder_name}/"):
                        state["selected_template"] = None
                    dlg.close()
                    refresh()
                except Exception as e:
                    ui.notify(str(e), type="negative")
            ui.button("Delete", on_click=confirm).classes(
                "bg-red-600 text-white"
            ).props("no-caps")
    dlg.open()


def _move_template_dialog(state: dict, refresh):
    """Dialog to move the selected template to a different folder."""
    selected_id = state.get("selected_template")
    if not selected_id:
        ui.notify("No template selected.", type="warning")
        return

    if "/" in selected_id:
        current_folder, name = selected_id.split("/", 1)
    else:
        current_folder, name = ROOT_FOLDER, selected_id

    folders = list_folders()
    target_folders = [f for f in folders if f != current_folder]

    if not target_folders:
        ui.notify("No other folders to move to.", type="info")
        return

    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label(f'Move template "{name}"')
        folder_select = ui.select(
            target_folders,
            value=target_folders[0],
            label="Target Folder",
        ).classes("w-full")
        folder_select.props("outlined")

        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def do_move():
                target = folder_select.value
                if not target:
                    ui.notify("Select a target folder.", type="warning")
                    return
                try:
                    move_template(selected_id, target)
                    state["selected_template"] = f"{target}/{name}"
                    ui.notify(f'Moved to "{target}".')
                    dlg.close()
                    refresh()
                except Exception as e:
                    ui.notify(f"Failed to move: {e}", type="negative")

            ui.button("Move", on_click=do_move).classes("bg-blue-600 text-white").props("no-caps")
    dlg.open()


def _add_new_template(state: dict, refresh):
    """Open dialog to enter folder + template name; create empty template and select it."""
    folders = list_folders()
    if not folders:
        # Ensure General exists by having at least one template or create folder on first template
        folders = [ROOT_FOLDER]
    with ui.dialog() as dlg, ui.card().classes("p-4 min-w-80"):
        ui.label("New template")
        folder_select = ui.select(
            folders,
            value=folders[0],
            label="Folder",
        ).classes("w-full")
        folder_select.props("outlined")
        name_input = ui.input("Template name").classes("w-full")
        name_input.props("outlined")
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dlg.close).props("flat no-caps")

            def create():
                name = (name_input.value or "").strip()
                folder = (folder_select.value or folders[0]).strip()
                if not name:
                    ui.notify("Enter a template name.", type="warning")
                    return
                try:
                    ui.notify("Template created.") # Moved before dlg.close()
                    template_id = f"{folder}/{name}"
                    write_template(template_id, "")
                    state["selected_template"] = template_id
                    dlg.close()
                    refresh()
                except Exception as e:
                    ui.notify(f"Failed to create: {e}", type="negative")
            ui.button("Create", on_click=create).classes("bg-blue-600 text-white").props("no-caps")
    dlg.open()


def _save_template(state: dict):
    selected = state.get("selected_template")
    editor = state.get("template_editor")
    if not selected or not editor:
        ui.notify("No template selected.", type="warning")
        return
    try:
        write_template(selected, editor.value or "")
        ui.notify("Template saved.")
    except Exception as e:
        ui.notify(f"Failed to save: {e}", type="negative")


def _send_mail(state: dict):
    outlook = OutlookClient()
    editor = state.get("template_editor")
    selected = state.get("selected_template") or "Draft"
    html = (editor.value or "") if editor else ""
    subject = (selected.split("/")[-1] if selected and "/" in selected else selected) or "No Subject"
    ok = outlook.create_draft(to="", subject=subject, body="", html_body=html or None)
    if ok:
        ui.notify("Draft created in Outlook.")
    else:
        ui.notify("Could not create draft. Is Outlook available?", type="negative")


def _delete_template(state: dict, refresh):
    selected = state.get("selected_template")
    if not selected:
        return
    display_name = selected.split("/")[-1] if "/" in selected else selected

    with ui.dialog() as dialog, ui.card().classes("p-4"):
        ui.label(f'Delete template "{display_name}"? This cannot be undone.')
        with ui.row().classes("mt-4 gap-2"):
            ui.button("Cancel", on_click=dialog.close).props("flat no-caps")
            def confirm():
                ui.notify("Template deleted.") # Moved before dialog.close()
                try:
                    delete_template(selected)
                    state["selected_template"] = None
                    dialog.close()
                    refresh()
                except Exception as e:
                    ui.notify(f"Failed to delete: {e}", type="negative")
            ui.button("Delete", on_click=confirm).classes(
                "bg-red-600 text-white"
            ).props("no-caps")
    dialog.open()


@ui.page("/")
def index():
    """Main app: sidebar + content area; content switches by nav."""
    ui.colors(primary="#2563eb", secondary="#64748b")
    ui.add_head_html(
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
    )

    state = {"section": "campaigns"}

    with ui.row().classes("w-full h-screen overflow-hidden").style(
        f"background-color: {NAVY};"
    ):
        # Left: sidebar (rebuilt on nav to update active highlight)
        sidebar_container = ui.column().classes("flex h-full")
        with sidebar_container:
            _build_sidebar(lambda s: None, state["section"])  # placeholder

        # Right: main content
        with ui.column().classes("flex-1 min-h-0 flex flex-col min-w-0"):
            content_holder = ui.column().classes(
                "flex-1 min-w-0 overflow-auto p-6 rounded-l-xl w-full"
            ).style(f"background-color: {SLATE_BG};")

    def render_content():
        content_holder.clear()
        with content_holder:
            if state["section"] == "campaigns":
                def refresh_campaigns():
                    content_holder.clear()
                    with content_holder:
                        if state["section"] == "campaigns":
                            _render_campaigns(state, refresh_campaigns)
                _render_campaigns(state, refresh_campaigns)
            elif state["section"] == "contact_lists":
                def refresh_contact_lists():
                    content_holder.clear()
                    with content_holder:
                        if state["section"] == "contact_lists":
                            _render_contact_lists(state, refresh_contact_lists)
                        else:
                            _placeholder_content(
                                "Contact Lists",
                                "Manage contact lists.",
                            )
                _render_contact_lists(state, refresh_contact_lists)
            else:
                # Template Library: two-panel list + editor; refresh re-renders this view
                def refresh_template_library():
                    content_holder.clear()
                    with content_holder:
                        if state["section"] == "template_library":
                            _render_template_library(state, refresh_template_library)
                        else:
                            _placeholder_content(
                                "Template Library",
                                "Manage email templates.",
                            )
                _render_template_library(state, refresh_template_library)

    def on_nav(section: str):
        state["section"] = section
        sidebar_container.clear()
        with sidebar_container:
            _build_sidebar(on_nav, section)
        render_content()

    # Wire sidebar to on_nav and show initial content
    sidebar_container.clear()
    with sidebar_container:
        _build_sidebar(on_nav, state["section"])
    render_content()


if __name__ in {"__main__", "__mp_main__"}:
    ui.run(
        title="SourceThatDeal",
        favicon="✉️",
        reload=True,
    )
