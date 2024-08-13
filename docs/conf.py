# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'SlySheets for Python'
copyright = '2023, Dunkyl 🔣🔣'
author = 'Dunkyl 🔣🔣'
release = '0.2.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    'myst_parser',
    'sphinxcontrib_trio',
    'sphinx_copybutton',
    'sphinxext.opengraph',
    "sphinx.ext.intersphinx",
    'sphinx.ext.autodoc',
    'sphinx.ext.duration',
    'sphinx.ext.autosummary',
    'sphinx.ext.napoleon',
    'sphinx.ext.viewcode',
    'sphinx.ext.coverage',
]

intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
    "aiohttp": ("https://docs.aiohttp.org/en/stable/", None),
}

napoleon_use_rtype = False
napoleon_numpy_docstring = False
napoleon_preprocess_types = True
napoleon_use_admonition_for_examples = True
napoleon_use_admonition_for_notes = True

autoclass_content = "both"
autosummary_generate = True
numpydoc_show_class_members = False

myst_heading_anchors = 3

autodoc_default_options = { # type: ignore
    "members": True,
    "inherited-members": False,
    "private-members": False,
    "show-inheritance": True,
    "undoc-members": True,
    "member-order": "bysource",
    "special-members": "__await__",
}

autodoc_member_order = 'bysource'
autodoc_typehints = "description"


templates_path = ['_templates']
exclude_patterns = []



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'furo'
html_static_path = ['_static']
html_title = 'SlySheets for Python'
html_theme_options = {
    "source_repository": "https://github.com/dunkyl/SlySheets-Python/",
    "source_branch": "main",
    "source_directory": "docs/",
}
html_favicon = '_static/sly logo py.svg'
html_logo = html_favicon

ogp_social_cards = {
    "enable": False
}
