/**
 * Rich WYSIWYG editor (Quasar QEditor) with full toolbar for email templates.
 * Toolbar: undo/redo, formatting, font sizes, lists, alignment, indent, link, headings.
 * Output is HTML suitable for Outlook HTMLBody (inline styles / simple structure).
 */
export default {
  template: `
    <q-editor
      ref="qRef"
      :id="id"
      v-model="inputValue"
      :toolbar="toolbar"
      min-height="12rem"
    >
      <template v-for="(_, slot) in $slots" v-slot:[slot]="slotProps">
        <slot :name="slot" v-bind="slotProps || {}" />
      </template>
    </q-editor>
  `,
  props: {
    value: String,
    id: String,
  },
  data() {
    return {
      inputValue: this.value,
      emitting: true,
      toolbar: [
        ['undo', 'redo'],
        ['bold', 'italic', 'underline', 'strike'],
        ['size-1', 'size-2', 'size-3', 'size-4', 'size-5', 'size-6', 'size-7'],
        ['unordered', 'ordered'],
        ['left', 'center', 'right', 'justify'],
        ['outdent', 'indent'],
        ['link'],
        ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'],
        ['quote', 'code'],
        ['removeFormat', 'hr'],
      ],
    };
  },
  beforeUnmount() {
    mounted_app.elements[this.$props.id.slice(1)].props.value = this.inputValue;
  },
  watch: {
    value(newValue) {
      this.emitting = false;
      this.inputValue = newValue;
      this.$nextTick(() => (this.emitting = true));
    },
    inputValue(newValue) {
      if (!this.emitting) return;
      this.$emit('update:value', newValue);
    },
  },
  methods: {
    updateValue() {
      this.inputValue = this.value;
    },
    insertAtCursor(text) {
      const editor = this.$refs.qRef;
      editor.focus();
      editor.runCmd('insertHTML', text);
    },
    setFontName(name) {
      const editor = this.$refs.qRef;
      editor.focus();
      editor.runCmd('fontName', name);
    },
    setFontSize(size) {
      const editor = this.$refs.qRef;
      editor.focus();
      editor.runCmd('fontSize', size);
    },
  },
};
