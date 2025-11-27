#!/usr/bin/env python
"""Django's command-line utility for administrative tasks."""
import os
import sys
from pathlib import Path


def main():
    """Run administrative tasks."""
    
    # BASE_DIR = pasta onde está o manage.py (attachflow_web)
    base_dir = Path(__file__).resolve().parent
    # ROOT_DIR = pasta acima (AttachFlow), onde vive o attachflow_core
    root_dir = base_dir.parent
    # garantir que o ROOT_DIR está no sys.path
    root_dir_str = str(root_dir)
    if root_dir_str not in sys.path:
        sys.path.insert(0, root_dir_str)
    
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'attachflow_web.settings')
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        raise ImportError(
            "Couldn't import Django. Are you sure it's installed and "
            "available on your PYTHONPATH environment variable? Did you "
            "forget to activate a virtual environment?"
        ) from exc
    execute_from_command_line(sys.argv)


if __name__ == '__main__':
    main()
