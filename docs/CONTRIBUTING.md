# Contributing to Civil 3D Geotable XML Report Generator

Thank you for your interest in contributing to this project! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Submission Guidelines](#submission-guidelines)
- [Testing](#testing)
- [Documentation](#documentation)

## Code of Conduct

This project follows a simple code of conduct:

- Be respectful and constructive
- Focus on technical merit
- Welcome newcomers and help them learn
- Acknowledge contributions from others

## How Can I Contribute?

### Reporting Bugs

Before creating a bug report:
1. Check existing issues to avoid duplicates
2. Verify the bug with the latest version
3. Collect relevant information (Civil 3D version, sample data, error messages)

Include in your bug report:
- **Description**: Clear description of the issue
- **Steps to Reproduce**: Detailed steps to replicate the bug
- **Expected Behavior**: What should happen
- **Actual Behavior**: What actually happens
- **Environment**: 
  - Civil 3D version
  - Dynamo version
  - Windows version
  - Sample data (if possible)
- **Screenshots**: If applicable
- **Error Messages**: Complete error text from Dynamo console

### Suggesting Enhancements

Enhancement suggestions are welcome! Include:
- **Use Case**: Describe the scenario where this would be useful
- **Proposed Solution**: How you envision it working
- **Alternatives**: Other approaches you've considered
- **Impact**: Who would benefit from this feature

### Code Contributions

We welcome code contributions! Areas where help is appreciated:

1. **Core Functionality**
   - Additional data extraction features
   - Performance optimizations
   - Error handling improvements

2. **Output Formats**
   - Additional export formats (CSV, JSON, GeoJSON)
   - Enhanced XML structure
   - Custom report templates

3. **Integration**
   - GIS integration utilities
   - BIM 360 connectivity
   - Database export capabilities

4. **Documentation**
   - Tutorial videos
   - Additional examples
   - Translations
   - API documentation

5. **Testing**
   - Unit tests
   - Integration tests
   - Sample data sets

## Development Setup

### Prerequisites

- Autodesk Civil 3D 2020 or later
- Python 3.7+
- Git
- Text editor or IDE (VS Code, PyCharm recommended)

### Getting Started

1. **Fork the Repository**
   ```bash
   # Fork on GitHub, then clone your fork
   git clone https://github.com/YOUR_USERNAME/xml-report-for-geotable.git
   cd xml-report-for-geotable
   ```

2. **Create a Branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bug-fix
   ```

3. **Set Up Development Environment**
   ```bash
   # Create virtual environment (optional, for standalone testing)
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   
   # Install development dependencies
   pip install -r requirements.txt
   ```

4. **Make Your Changes**
   - Edit files as needed
   - Follow coding standards (see below)
   - Test your changes thoroughly

5. **Commit Your Changes**
   ```bash
   git add .
   git commit -m "Description of your changes"
   ```

6. **Push to Your Fork**
   ```bash
   git push origin feature/your-feature-name
   ```

7. **Create Pull Request**
   - Go to GitHub and create a pull request
   - Fill out the pull request template
   - Link any related issues

## Coding Standards

### Python Code Style

Follow **PEP 8** guidelines:

```python
# Good: Clear function names, docstrings, type hints
def extract_alignment_data(alignment_name: str, interval: float = 10.0) -> dict:
    """
    Extract geotable data from a Civil 3D alignment.
    
    Parameters:
    - alignment_name: Name of the alignment to extract
    - interval: Station sampling interval in drawing units
    
    Returns:
    - Dictionary containing extracted geotable data
    """
    # Implementation
    pass

# Bad: Unclear names, no documentation
def ead(a, i=10):
    # What does this do?
    pass
```

### Specific Guidelines

**Naming Conventions:**
- Classes: `PascalCase` (e.g., `GeoTableDataExtractor`)
- Functions/Methods: `snake_case` (e.g., `extract_alignment_data`)
- Constants: `UPPER_SNAKE_CASE` (e.g., `DEFAULT_INTERVAL`)
- Private methods: `_leading_underscore` (e.g., `_validate_input`)

**Documentation:**
- All public functions must have docstrings
- Include parameter descriptions and return types
- Add inline comments for complex logic
- Update README.md for new features

**Error Handling:**
```python
# Good: Specific error handling with helpful messages
try:
    alignment = get_alignment(name)
except ObjectNotFoundException:
    return f"Error: Alignment '{name}' not found in drawing"
except Exception as e:
    return f"Unexpected error: {str(e)}"

# Bad: Bare except that hides errors
try:
    alignment = get_alignment(name)
except:
    pass
```

**Code Organization:**
- Keep functions focused (single responsibility)
- Limit functions to ~50 lines when possible
- Group related functions in classes
- Use meaningful variable names

### XML and JSON

**XML Structure:**
- Follow the existing schema in `geotable_schema.xsd`
- Use clear, descriptive element names
- Include appropriate attributes
- Validate against schema

**JSON Configuration:**
- Use consistent formatting (2-space indentation)
- Include comments explaining options
- Provide sensible defaults
- Document all configuration keys

## Submission Guidelines

### Pull Request Process

1. **Update Documentation**
   - Update README.md if adding features
   - Add examples to example_usage.py
   - Update CHANGELOG.md
   - Document configuration changes in config.json

2. **Test Your Changes**
   - Test with Civil 3D (multiple versions if possible)
   - Test with various alignment types
   - Test error conditions
   - Verify XML output is valid

3. **Code Review**
   - Address review comments promptly
   - Be open to suggestions
   - Explain your design decisions

4. **Merge Criteria**
   - All tests pass
   - Documentation updated
   - Code follows style guidelines
   - No breaking changes (unless discussed)
   - Approved by maintainers

### Commit Message Format

Use clear, descriptive commit messages:

```
Good examples:
- "Add support for vertical curve extraction"
- "Fix station interval validation bug"
- "Update XML schema for cant data"
- "Improve error messages in extractor"

Less helpful:
- "Update code"
- "Fix bug"
- "Changes"
```

## Testing

### Testing Your Code

**Manual Testing in Civil 3D:**
1. Open Civil 3D with test drawing
2. Load modified scripts in Dynamo
3. Test with various alignments:
   - Simple straight alignments
   - Complex curve combinations
   - Alignments with superelevation
   - Very short/long alignments
4. Verify XML output
5. Check edge cases and error conditions

**Test Cases to Cover:**
- Empty/null inputs
- Invalid alignment names
- Zero or negative intervals
- Very large intervals
- Alignments with no geometry
- Corrupted alignment data
- File I/O errors

**Sample Test Data:**
Create or provide sample Civil 3D drawings with:
- Various alignment geometries
- Different coordinate systems
- Rail-specific elements
- Edge cases

### Automated Testing

If adding Python unit tests:

```python
# Example test structure
import unittest

class TestGeoTableExtractor(unittest.TestCase):
    
    def test_interval_validation(self):
        """Test that invalid intervals are rejected"""
        # Test code here
        pass
    
    def test_coordinate_extraction(self):
        """Test coordinate extraction accuracy"""
        # Test code here
        pass
```

Run tests:
```bash
python -m pytest tests/
```

## Documentation

### Documentation Standards

**Code Documentation:**
- All public APIs must have docstrings
- Include usage examples in docstrings
- Document exceptions that may be raised
- Explain complex algorithms

**User Documentation:**
- Update README.md for new features
- Add entries to CHANGELOG.md
- Create examples in example_usage.py
- Consider adding tutorial content

**API Documentation:**
- Document all parameters and return values
- Provide usage examples
- Note version compatibility
- List dependencies

### Documentation Templates

**Function Documentation:**
```python
def function_name(param1: type1, param2: type2 = default) -> return_type:
    """
    Brief description of what the function does.
    
    More detailed explanation if needed, including:
    - Algorithm description
    - Important notes
    - Usage context
    
    Parameters:
    - param1: Description of param1
    - param2: Description of param2 (default: default_value)
    
    Returns:
    - Description of return value
    
    Raises:
    - ExceptionType: When this exception occurs
    
    Example:
        >>> result = function_name("value", 42)
        >>> print(result)
        expected_output
    """
```

## Questions?

If you have questions about contributing:

1. Check existing documentation (README.md, wiki)
2. Search existing issues
3. Open a new issue with your question
4. Tag it appropriately

## Recognition

Contributors will be:
- Listed in CHANGELOG.md
- Acknowledged in release notes
- Added to contributors list (if desired)

Thank you for contributing to make this tool better for everyone!



