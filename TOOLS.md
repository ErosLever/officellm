# Tool Reference

All 129 MCP tools exposed by this project, grouped by host application and category. Every tool (except `office_get_active_apps` and a handful of server-side cross-cutting tools) requires an `instanceId` obtained from `office_get_active_apps`.

## Shared / Cross-cutting

| Tool | Description |
|---|---|
| `office_get_active_apps` | Lists all registered Office instances (call this first) |
| `office_get_document_context` | Rich summary of app, document metadata, and selection state |
| `office_get_document_stats` | Quantifiable metrics: word/page/slide/cell/item counts |
| `office_batch_call` | Executes up to 10 tool calls in parallel |
| `office_suggest_tools` | Suggests relevant tools for the current host/category |
| `office_export_document` | Exports the document as PDF or native format (base64) |

## PowerPoint

### Read
| Tool | Description |
|---|---|
| `powerpoint_get_deck_outline` | Full slide deck outline with shapes per slide |
| `powerpoint_get_slide` | All shapes on a slide with position, size, text, styling |
| `powerpoint_get_slide_image` | Renders a slide as a PNG (base64) |
| `powerpoint_get_shape_image` | Renders a single shape as a PNG (base64) |
| `powerpoint_get_table` | Reads all cell text from a table shape |
| `powerpoint_get_selection` | Current user selection (text or shapes) |
| `powerpoint_get_speaker_notes` | Speaker notes for one slide or a range |

### Write
| Tool | Description |
|---|---|
| `powerpoint_update_shape_text` | Updates a shape's text content |
| `powerpoint_update_shape_properties` | Updates position, size, rotation, font of a shape |
| `powerpoint_update_speaker_notes` | Replaces speaker notes for a slide |

### Shape CRUD
| Tool | Description |
|---|---|
| `powerpoint_add_textbox` | Creates a new text box |
| `powerpoint_add_image` | Inserts an image from base64 data |
| `powerpoint_add_table` | Creates a new table shape |
| `powerpoint_delete_shape` | Deletes a shape (irreversible) |

### Slide Management
| Tool | Description |
|---|---|
| `powerpoint_add_slide` | Inserts a new blank slide |
| `powerpoint_delete_slide` | Deletes a slide (irreversible) |
| `powerpoint_move_slide` | Moves a slide to a new position |

### Tags & Metadata
| Tool | Description |
|---|---|
| `powerpoint_get_tags` | Reads key-value tags on a presentation/slide/shape |
| `powerpoint_set_tag` | Sets a key-value tag for audience filtering |
| `powerpoint_delete_slides_by_tag` | Deletes all slides matching a tag key/value |

### Shape Formatting
| Tool | Description |
|---|---|
| `powerpoint_set_shape_fill` | Sets fill color, transparency, or image fill |
| `powerpoint_set_shape_line` | Sets border color, width, and style |
| `powerpoint_set_shape_rotation` | Rotates a shape by degrees |

### Geometric Shapes & Lines
| Tool | Description |
|---|---|
| `powerpoint_add_geometric_shape` | Adds a rectangle, oval, arrow, star, etc. |
| `powerpoint_add_line` | Adds a straight, elbow, or curved connector |

### Slide Merge
| Tool | Description |
|---|---|
| `powerpoint_insert_slides_from_file` | Inserts slides from another PPTX (base64) |

### Layouts & Theme
| Tool | Description |
|---|---|
| `powerpoint_get_layouts` | Lists slide layouts from the slide master |
| `powerpoint_get_theme_colors` | Returns the 10 theme colors |
| `powerpoint_group_shapes` | Groups multiple shapes into one |
| `powerpoint_ungroup_shape` | Ungroups a grouped shape |

## Word

### Read
| Tool | Description |
|---|---|
| `word_get_outline` | Document outline (headings, levels, styles) |
| `word_get_paragraphs` | Paragraphs from the body, with range filtering |
| `word_get_selection` | Currently selected text with paragraph context |
| `word_search` | Searches text; returns paragraphIndex + snippet per match |

### Write
| Tool | Description |
|---|---|
| `word_replace_text` | Replaces text within a specific paragraph |
| `word_insert_text` | Inserts text at end / before / after a paragraph |
| `word_replace_selection` | Replaces the current selection with new text |
| `word_find_replace` | Find & replace with wildcards, case, whole-word, scoping |
| `word_delete_paragraph` | Deletes a paragraph by index |

### Tracked Changes
| Tool | Description |
|---|---|
| `word_get_tracked_changes` | Current tracking mode and pending change count |
| `word_accept_all_changes` | Accepts all pending tracked changes |
| `word_reject_all_changes` | Rejects all pending tracked changes |

### Tables
| Tool | Description |
|---|---|
| `word_get_tables` | All tables with row/column counts and cell text |
| `word_insert_table` | Inserts a table with optional header row |
| `word_update_table_cell` | Updates a single cell's text |
| `word_add_table_rows` | Adds rows (matrix or object array with headers) |
| `word_delete_table_row` | Deletes a row by index |
| `word_delete_table_column` | Deletes a column, auto-adjusting outer borders |
| `word_add_table_column` | Adds a column, auto-adjusting outer borders |
| `word_merge_table_cells` | Merges a rectangular range of cells |
| `word_split_table_cell` | Splits a cell into a row×column grid |
| `word_copy_table_structure` | Clones a table's structure (borders, widths, padding) |
| `word_set_table_format` | Sets header rows, style, alignment, padding, widths, per-column overrides |

### Sections & Headers/Footers
| Tool | Description |
|---|---|
| `word_get_sections` | Document sections with page layout and header/footer config |
| `word_get_headers_footers` | Header/footer content per section and variant |
| `word_set_header_footer` | Sets header or footer text for a section |

### Lists
| Tool | Description |
|---|---|
| `word_insert_list` | Inserts a bulleted or numbered list |

### Styles & Formatting
| Tool | Description |
|---|---|
| `word_apply_style` | Applies a named paragraph or character style |
| `word_get_styles` | All defined styles with font/paragraph properties |
| `word_modify_style` | Modifies an existing style's definition document-wide |
| `word_create_style` | Creates a new named style, optionally inheriting a base |
| `word_create_and_remap_style` | Clones + overrides a style and remaps existing paragraphs to it |
| `word_get_formatting` | Paragraph + font formatting for one or more paragraphs |
| `word_set_formatting` | Applies paragraph/font formatting to paragraphs or selection |

### Bookmarks, Fields & Hyperlinks
| Tool | Description |
|---|---|
| `word_get_bookmarks` | All bookmarks with names and ranges |
| `word_insert_bookmark` | Creates/updates a bookmark at a paragraph range |
| `word_delete_bookmark` | Deletes a bookmark (keeps the text) |
| `word_goto_bookmark` | Navigates to and selects a bookmark |
| `word_get_hyperlinks` | All hyperlinks with text and URLs |
| `word_insert_hyperlink` | Inserts a hyperlink at a location |
| `word_insert_field` | Inserts a field (TOC, page number, date, etc.) |

### Footnotes, Endnotes & Fields
| Tool | Description |
|---|---|
| `word_insert_footnote` | Inserts a footnote after a paragraph |
| `word_insert_endnote` | Inserts an endnote after a paragraph |

### Document Properties & Content Controls 
| Tool | Description |
|---|---|
| `word_get_properties` | Built-in document properties (title, author, etc.) |
| `word_set_properties` | Updates document properties |
| `word_get_content_controls` | All content controls with titles, tags, content |
| `word_insert_content_control` | Wraps a paragraph range in a content control |

### Comments
| Tool | Description |
|---|---|
| `word_get_comments` | All comment threads with replies and anchor text |
| `word_add_comment` | Adds a comment anchored to text, paragraphs, or selection |
| `word_edit_comment` | Edits a top-level comment's text |
| `word_resolve_comment` | Resolves or reopens a comment thread |
| `word_delete_comment` | Deletes a comment thread and its replies |
| `word_reply_to_comment` | Adds a reply to a comment thread |
| `word_edit_reply` | Edits a specific reply's text |
| `word_delete_reply` | Deletes a single reply |

### Images
| Tool | Description |
|---|---|
| `word_get_image` | Returns an inline image as a base64 data URL for visual analysis |
| `word_insert_image` | Inserts an inline image from base64 data |

## Excel

### Read
| Tool | Description |
|---|---|
| `excel_get_workbook_map` | Sheet names, used ranges, tables, named ranges |
| `excel_read_range` | Reads values/formulas/number formats from a range |

### Write
| Tool | Description |
|---|---|
| `excel_write_range` | Writes a 2D array of values to a range |
| `excel_write_formula` | Writes a formula to a cell or range |
| `excel_create_table` | Creates a formatted Excel table (ListObject) from a range |

### Sheet Management
| Tool | Description |
|---|---|
| `excel_add_sheet` | Adds a new worksheet |
| `excel_delete_sheet` | Deletes a worksheet by name |
| `excel_rename_sheet` | Renames a worksheet |

### Sort & Filter
| Tool | Description |
|---|---|
| `excel_sort_range` | Multi-column sort with ascending/descending order |
| `excel_filter_range` | Applies or clears autofilter with value-based criteria |

### Charts
| Tool | Description |
|---|---|
| `excel_create_chart` | Creates a chart (Column, Bar, Line, Pie, etc.) from a range |
| `excel_get_charts` | Lists charts with type, title, data range, position |

### Formatting
| Tool | Description |
|---|---|
| `excel_format_range` | Font, fill, borders, alignment, number format |
| `excel_apply_conditional_formatting` | Cell value rules, data bars, color scales, icon sets |

### Pivot Table
| Tool | Description |
|---|---|
| `excel_create_pivottable` | Creates a pivot table with row/column/value fields |

### Navigation
| Tool | Description |
|---|---|
| `excel_freeze_panes` | Freezes rows/columns for scroll navigation |
| `excel_get_named_ranges` | Lists named ranges with addresses and scopes |
| `excel_add_named_range` | Creates or updates a named range |

### Data Validation
| Tool | Description |
|---|---|
| `excel_add_data_validation` | Dropdown lists, number/date/text constraints, custom formulas |
| `excel_remove_data_validation` | Removes data validation from a range |

### Protection & Page Layout
| Tool | Description |
|---|---|
| `excel_protect_sheet` | Protects a worksheet with password and granular permissions |
| `excel_unprotect_sheet` | Removes worksheet protection |
| `excel_set_page_layout` | Orientation, paper size, margins, print area/title rows |
| `excel_get_page_layout` | Returns current page layout settings |

## Outlook

### Read
| Tool | Description |
|---|---|
| `outlook_get_current_item` | Metadata and body of the selected item |
| `outlook_summarize_thread` | Structured summary of an email thread |

### Write (never auto-sends)
| Tool | Description |
|---|---|
| `outlook_draft_reply` | Creates a draft reply in the Drafts folder |
| `outlook_apply_category` | Applies a color-coded category to selected email(s) |
| `outlook_send_message` | Sends a draft — requires explicit confirmation token |

### Extended (Office.js only)
| Tool | Description |
|---|---|
| `outlook_get_user_profile` | Display name, email, timezone of current user |
| `outlook_get_master_categories` | Mailbox's master category list |
| `outlook_create_category` | Adds a new category to the master list |
| `outlook_remove_categories` | Removes categories from the current item |
| `outlook_display_new_message` | Opens a compose form (user sends manually) |
| `outlook_display_new_appointment` | Opens a new appointment form (user saves manually) |
| `outlook_get_attachments` | Attachment metadata (name, size, type, inline status) |
