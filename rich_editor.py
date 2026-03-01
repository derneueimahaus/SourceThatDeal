"""
Rich WYSIWYG editor for email templates (Quasar QEditor with full toolbar).
Output is HTML suitable for Outlook HTMLBody. Use in Template Library instead of ui.editor.
"""

from typing import Any

from nicegui.elements.mixins.disableable_element import DisableableElement
from nicegui.elements.mixins.value_element import ValueElement
from nicegui.defaults import DEFAULT_PROP, resolve_defaults
from nicegui.events import Handler, ValueChangeEventArguments


class RichEditor(ValueElement, DisableableElement, component="rich_editor.js", default_classes="nicegui-editor"):
    """WYSIWYG editor with full toolbar (font size, bold, italic, lists, link, alignment, etc.) for email templates."""

    VALUE_PROP: str = "value"
    LOOPBACK = False

    @resolve_defaults
    def __init__(
        self,
        *,
        placeholder: str | None = DEFAULT_PROP | None,
        value: str = DEFAULT_PROP | "",
        on_change: Handler[ValueChangeEventArguments] | None = None,
        toolbar: list[list[str]] | None = None,
    ) -> None:
        super().__init__(value=value, on_value_change=on_change)
        self._props.set_optional("placeholder", placeholder)
        self._props.set_optional("toolbar", toolbar)

    def _handle_value_change(self, value: Any) -> None:
        super()._handle_value_change(value)
        if self._send_update_on_value_change:
            self.run_method("updateValue")

    def insert_at_cursor(self, text: str) -> None:
        """Insert text (e.g. a placeholder like '{{First Name}}') at the current cursor position."""
        self.run_method("insertAtCursor", text)

    def set_font_name(self, name: str) -> None:
        """Apply a font family to the selected text."""
        self.run_method("setFontName", name)

    def set_font_size(self, size: int) -> None:
        """Apply a font size (1-7) to the selected text."""
        self.run_method("setFontSize", size)
