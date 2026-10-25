# -*- coding: utf-8 -*-
from .dotx_renderer import render_dotx, render_styles_into, save_as_dotx, apply_page_setup
from .apply import apply_profile_offline, resolve_assignment, strip_direct_formatting

__all__ = ["render_dotx", "render_styles_into", "save_as_dotx", "apply_page_setup",
           "apply_profile_offline", "resolve_assignment", "strip_direct_formatting"]
