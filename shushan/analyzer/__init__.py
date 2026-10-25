# -*- coding: utf-8 -*-
from .docx_analyzer import analyze_document, AnalysisResult, ParaInfo
from .effective import format_signature, read_page_setup

__all__ = ["analyze_document", "AnalysisResult", "ParaInfo",
           "format_signature", "read_page_setup"]
