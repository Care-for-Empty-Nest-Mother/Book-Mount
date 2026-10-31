# -*- coding: utf-8 -*-
from .dotx_renderer import render_dotx, render_styles_into, save_as_dotx, apply_page_setup
from .apply import apply_profile_offline, resolve_assignment, strip_direct_formatting
from .page_furniture import (apply_headers_footers, apply_toc_updates,
                             insert_toc_field, set_update_fields,
                             restyle_table_cells, paragraph_section_map)

__all__ = ["render_dotx", "render_styles_into", "save_as_dotx", "apply_page_setup",
           "apply_profile_offline", "resolve_assignment", "strip_direct_formatting",
           "apply_headers_footers", "apply_toc_updates", "insert_toc_field",
           "set_update_fields", "restyle_table_cells", "paragraph_section_map"]
