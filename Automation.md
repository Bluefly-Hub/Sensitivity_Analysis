# Development Guidelines & Preferences

## Brad's Coding Standards & Preferences

### Version Control & Git
- **Always use `.gitignore`** to exclude build artifacts and virtual environments
- Ignore folders: `.venv/`, `venv/`, `env/`, `build/`, `dist/`, `__pycache__/`, `tmp_*/`, `.pytest_cache/`
- Ignore files: `*.spec`, `*.pyc`, `*.pyo`, `*.pyd`, `.DS_Store`, `Thumbs.db`, IDE configs
- Commit only source code, configuration files, and documentation
- Never commit compiled executables, build artifacts, or sensitive data to the repository
- Use meaningful commit messages that describe what changed and why

### Dependency Management
- **Always maintain a `requirements.txt`** file for Python projects
- Keep it minimal - only include packages actually imported in the code
- Remove transitive dependencies (packages installed automatically by other packages)
- Update dependencies periodically but test thoroughly after updates
- Consider using version pinning for production applications

### Application Updates & Versioning
- **Use GitHub releases for distribution** of compiled applications
- Implement auto-updater functionality when distributing to end users:
  - Check GitHub API on application startup
  - Prompt users with version information
  - Download and self-replace executable automatically
  - Restart application after update
- Maintain version tracking in `version.py` or similar
- Use semantic versioning (e.g., `1.0.0`, `1.0.1`, `1.1.0`)
  - MAJOR: Breaking changes
  - MINOR: New features, backwards compatible
  - PATCH: Bug fixes

### Build Process (Python Projects)
- Use PyInstaller for creating Windows executables
- Always use `--onefile` flag for single executable distribution
- Use `--windowed` flag for GUI applications (no console window)
- Use `--console` or omit flag for CLI applications
- Standard build command: `pyinstaller --onefile --windowed --name "AppName" main.py`
- Distribute only the final `.exe` file from the `dist/` folder
- Test executables on clean machines without Python installed

### Code Organization
- Keep entry point simple and clear (`main.py`, `app.py`, or similar)
- Separate concerns into dedicated modules:
  - GUI/UI logic in separate files
  - Business logic in dedicated modules
  - Utility functions in helper modules
  - Configuration in config files or modules
- Use descriptive file and module names
- Keep functions focused on a single responsibility
- Add docstrings to modules, classes, and complex functions

### AI Agent Preferences
- **No approval required for package installations** - Install dependencies automatically
- **No approval required for file operations** - Create, edit, and delete files as needed
- **No approval required for terminal commands** - Run builds, tests, and Git operations freely
- **Prefer action over asking** - Implement solutions directly rather than just suggesting
- **Make reasonable assumptions** - When details are unclear, infer from context and proceed
- **Be proactive** - Anticipate needs (e.g., create .gitignore, update requirements.txt)
- **Fix issues directly** - Don't just identify problems, solve them
- **Default to implementing** - Build working solutions, not just explanations

### Platform Considerations
**Windows (Primary Environment):**
- Be aware of Windows long path limitations (260 character limit)
- Use `.gitignore` to avoid committing files with long paths
- Enabling Windows long paths is safe (no performance impact)
- Use PowerShell-compatible commands in terminal operations
- Test path separators (use `os.path.join()` for cross-platform compatibility)

**General:**
- Consider cross-platform compatibility when possible
- Use `pathlib` for modern path handling
- Avoid hardcoding file paths

### Error Handling
- Provide user-friendly error messages, not just stack traces
- Handle common failure modes gracefully (network, file I/O, permissions)
- Don't crash on missing dependencies - attempt to install or guide user
- Log errors appropriately (console for debugging, files for production)
- Validate user inputs before processing
- Use try-except blocks for external operations (API calls, file operations)

### Distribution Workflow (For Compiled Apps)
1. Update version number in version tracking file
2. Rebuild executable with PyInstaller
3. Test the executable locally on a clean machine if possible
4. Create GitHub Release with semantic version tag (e.g., `v1.0.0`)
5. Upload the executable to the release
6. Update release notes with changes and features
7. Users with auto-updater receive notifications automatically

### General Best Practices
- Keep virtual environments isolated and never commit them
- Use meaningful variable, function, and file names
- Write self-documenting code with clear intent
- Add comments for complex logic, not obvious statements
- Follow PEP 8 style guidelines for Python
- Test edge cases and error conditions
- Keep functions small and focused
- Don't repeat yourself (DRY principle)
- Maintain backwards compatibility when feasible
- Document breaking changes clearly

## Quick Reference Commands (Python/Windows)

### Setup New Python Environment
```powershell
# Create virtual environment
python -m venv .venv

# Activate (PowerShell)
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

### Build Executable (PyInstaller)
```powershell
# GUI Application (no console)
pyinstaller --onefile --windowed --name "AppName" main.py

# CLI Application (with console)
pyinstaller --onefile --name "AppName" main.py
```

### Update Dependencies
```powershell
# Export all installed packages
pip freeze > requirements.txt

# Then manually clean to keep only direct dependencies
# Remove packages that are auto-installed by others
```

### Git Workflow
```powershell
# Check status
git status

# Add files
git add .

# Commit with message
git commit -m "Descriptive message of changes"

# Push to remote
git push origin main

# Create and push tag for release
git tag v1.0.0
git push origin v1.0.0
```

### Testing
```powershell
# Run Python script
python main.py

# Run with specific Python version
python3.13 main.py

# Run tests (if using pytest)
pytest
```

## Project-Specific Notes
- Work is primarily done in Windows environments
- PowerShell is the primary shell
- Python projects should use virtual environments
- Prefer single-file executables for distribution
- GitHub is used for version control and releases
