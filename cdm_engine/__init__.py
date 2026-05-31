# cdm_engine/__init__.py
"""CDM Engine — unified clinical data management toolkit for Excel workflows.

Usage:
    from cdm_engine import standardise_terms, generate_toc
    from cdm_engine import merge_workbooks, convert_lab_units
"""
__version__ = "3.0.0"

from cdm_engine.engine import (
    # Text
    standardise_terms,
    # Duplicates
    highlight_duplicates,
    # Blank rows
    insert_blank_every_n,
    insert_blank_between_groups,
    delete_blank_rows,
    # TOC
    generate_toc,
    generate_sas_derive_toc,
    # Sheet management
    batch_rename_sheets,
    list_sheet_names,
    delete_hidden_sheets,
    create_sheets_from_list,
    delete_sheets,
    # Track changes
    track_changes,
    changes_report,
    # Strikethrough
    delete_strikethrough_rows,
    clear_strikethrough,
    # Medical calculators
    # Special chars
    scan_special_chars,
    # File listing
    generate_file_listing,
    generate_folder_tree,
    # Dates
    add_days_to_column,
    calc_date_diff,
    # Lab units
    convert_lab_units,
    # Medical calculators
    calc_cockcroft_gault,
    calc_ckd_epi,
    # Utilities
    # Sheet split
    split_sheet_by_column,
    # Workbook merge
    merge_workbooks,
    # File-based wrappers (for Web API)
    merge_workbooks_from_uploads,
    split_sheet_file,
    add_days_to_column_file,
    convert_lab_units_file,
    list_directory,
    # Utilities
    apply_header_style,
    auto_width,
)
