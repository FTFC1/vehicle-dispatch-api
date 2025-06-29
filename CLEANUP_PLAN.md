# Project Cleanup Plan - V4 Iteration

## Current State Assessment
- **Working Files**: `api_only.py` (updated with auto-detection), `simpler_processor.py` (core logic)
- **Legacy Files**: Multiple Flask iterations, analyzers, test files
- **Bloat**: 123KB log file, empty directories, abandoned experiments

## V4 Core Requirements
- Flask backend API (clean, single file)
- v0.dev frontend integration
- Excel processing (11 brands, VIN splitting)
- Deployment ready

## Cleanup Actions

### Keep (Essential for V4)
- `api_only.py` - Main Flask API (already updated)
- `simpler_processor.py` - Core processing logic
- `requirements.txt` - Dependencies
- `Files/` directory - Test data
- `README.md` - Documentation

### Archive (Legacy from previous iterations)
- `app.py`, `app_fixed.py`, `fast_app.py`, `simple_app.py` - Old Flask attempts
- `processor.py`, `custom_format_processor.py` - Legacy processors
- All analyzer/fixer files - Experimental code
- Test files (`test_*.py`, `minimal_test.py`)
- `templates/`, `static/` directories - Old UI attempts

### Delete (Clutter)
- `vininspector.log` (123KB of logs)
- `test_output.xlsx` (failed test artifact)
- `__pycache__/` directories
- `grok chat` (conversation backup)
- Empty directories

### Clean Structure for V4
```
operations-flask-app/
├── api_only.py          # Main Flask API
├── simpler_processor.py # Core processing logic  
├── requirements.txt     # Dependencies
├── README.md           # Documentation
├── Files/              # Test data
└── archive/            # Legacy files
```

## Next Steps
1. Archive legacy files
2. Test clean API
3. Build v0.dev frontend
4. Deploy backend 