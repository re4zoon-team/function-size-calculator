# Function Size Calculator

A Python tool that scans git repositories to find the largest functions in Java, Node.js, and Python codebases. The results are exported to an Excel (XLSX) or JSON file with each repository on a separate tab.

## Features

- Scans multiple git repositories (local or remote)
- **Parallel processing** for efficient scanning of multiple repositories
- Supports Node.js (JavaScript, TypeScript), Java, and Python
- **Memory-efficient streaming** handles very large files without loading entire files into memory
- **Multiple output formats**: Excel (XLSX) and JSON
- **Configurable number of top functions** to report (default: 5)
- **Minimum function size filter** to exclude trivial functions
- **Summary statistics** for each repository
- **Git clone timeout** to prevent hanging on problematic repositories
- Exports results to Excel format with:
  - Each repository on a separate tab
  - Function name, file path, line numbers, and size
  - Summary statistics (total functions, average size, largest/smallest)
  - Formatted headers and auto-sized columns
- JSON export option for programmatic consumption
- Automatic cleanup of temporary cloned repositories

## Installation

1. Clone this repository:

```bash
git clone https://github.com/Re4zOon/function-size-calculator.git
cd function-size-calculator
```

1. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Requirements

- Python 3.14 or higher
- Git (for cloning remote repositories)
- openpyxl (automatically installed from requirements.txt)

## Testing

The repository includes a comprehensive test suite to ensure code quality and prevent regressions.

### Running Tests

Run all tests (pytest is configured for colorful, verbose output):

```bash
pytest
```

Run specific test class:

```bash
pytest tests/test_function_size_calculator.py -k TestJavaScriptParser
```

Run a specific test:

```bash
pytest tests/test_function_size_calculator.py -k test_parse_javascript_file
```

### Test Coverage

The test suite includes:

- **Unit tests** for FunctionInfo class
- **Parser tests** for JavaScript/TypeScript and Java parsers
- **Excel writer tests** for output generation
- **Integration tests** for repository scanning
- **CLI tests** for command-line argument parsing

Test fixtures are located in `tests/fixtures/` and include sample JavaScript, TypeScript, and Java files.

## Usage

### Basic Usage

Scan one or more repositories:

```bash
python function_size_calculator.py <repository-path-or-url> [<repository-path-or-url> ...]
```

### Examples

1. Scan a remote repository:

```bash
python function_size_calculator.py https://github.com/user/repo.git
```

1. Scan multiple repositories:

```bash
python function_size_calculator.py https://github.com/user/repo1.git https://github.com/user/repo2.git
```

1. Scan local repositories:

```bash
python function_size_calculator.py /path/to/local/repo1 /path/to/local/repo2
```

1. Scan repositories from an input file:

```bash
# Create a file with repository URLs (one per line)
cat > repos.txt << EOF
https://github.com/user/repo1.git
https://github.com/user/repo2.git
/path/to/local/repo3
# Comments are supported
https://github.com/user/repo4.git
EOF

# Scan all repositories from the file
python function_size_calculator.py -i repos.txt
```

1. Extract repository URLs from a YAML file:

```bash
# If you have a YAML file with repository information:
# - name: spring-boot-template
#   url: https://gitlab.local/services/spring-boot-template
#   product: Devops
#   type: springboot

# Use yaml_to_repos.py to extract URLs to repos.txt
python yaml_to_repos.py repos.yaml -o repos.txt

# Then scan using the generated repos.txt
python function_size_calculator.py -i repos.txt
```

1. Specify custom output file:

```bash
python function_size_calculator.py -o my_results.xlsx https://github.com/user/repo.git
```

1. Mix input file and command-line repositories:

```bash
python function_size_calculator.py -i repos.txt https://github.com/user/extra-repo.git
```

1. Adjust parallel processing (default is 4 parallel jobs):

```bash
python function_size_calculator.py -i repos.txt -j 8  # Use 8 parallel jobs
```

1. Report top 10 functions instead of default 5:

```bash
python function_size_calculator.py -i repos.txt -n 10
```

1. Filter out small functions (e.g., exclude functions smaller than 10 lines):

```bash
python function_size_calculator.py -i repos.txt -m 10
```

1. Combine multiple options:

```bash
python function_size_calculator.py -i repos.txt -n 20 -m 5 -j 8 -o detailed_analysis.xlsx
```

1. Generate JSON output instead of Excel:

```bash
python function_size_calculator.py -i repos.txt -o results.json
```

1. Explicitly specify output format:

```bash
python function_size_calculator.py -i repos.txt -f json -o results.json
```

### Command-Line Options

- `repositories`: One or more git repository URLs or local paths (optional if using -i)
- `-i`, `--input-file`: File containing list of repository URLs/paths (one per line, comments with # are supported)
- `-o`, `--output`: Output file name (default: `function_sizes.xlsx`). Use .json extension for JSON format.
- `-f`, `--format`: Output format - `xlsx`, `json`, or `auto` (default: auto - detect from file extension)
- `-j`, `--jobs`: Number of parallel jobs for scanning repositories (default: 4)
- `-n`, `--top-n`: Number of top largest functions to report per repository (default: 5)
- `-m`, `--min-size`: Minimum function size in lines to include (default: 1)
- `-h`, `--help`: Show help message

### YAML to repos.txt Converter (yaml_to_repos.py)

If you have repository information in a YAML file, use `yaml_to_repos.py` to extract the URLs:

```bash
python yaml_to_repos.py repos.yaml -o repos.txt
```

**Expected YAML structure:**

```yaml
- name: spring-boot-template
  url: https://gitlab.local/services/spring-boot-template
  product: Devops
  type: springboot
- name: spring-boot
  url: https://gitlab.local/services/spring-boot
  product: Devops
  type: springboot
```

**Options:**
- `yaml_file`: Path to the YAML file containing repository information (required)
- `-o`, `--output`: Output file name (default: `repos.txt`)

## Output Formats

### Excel (XLSX)

The default output format. Generates an Excel file with the following structure:

- **Each repository gets its own tab** named after the repository
- **Columns in each tab:**
  - Rank: Position in top N (1-N based on --top-n parameter)
  - Function Name: Name of the function/method
  - File Path: Relative path to the file containing the function
  - Start Line: Line number where the function starts
  - End Line: Line number where the function ends
  - Lines of Code: Total lines in the function
- **Summary Statistics:**
  - Total Functions Found
  - Average Function Size
  - Largest Function
  - Smallest Function

### JSON

An alternative output format for programmatic consumption. Structure:

```json
{
  "repository-name": {
    "summary": {
      "total_functions": 100,
      "average_size": 15.5,
      "largest_function_size": 250,
      "smallest_function_size": 3
    },
    "top_functions": [
      {
        "name": "functionName",
        "file_path": "path/to/file.js",
        "start_line": 10,
        "end_line": 50,
        "size": 41
      }
    ]
  }
}
```

To use JSON format, either:

- Use the `.json` extension: `-o results.json`
- Explicitly specify format: `-f json -o results.json`

## Supported Languages

### Node.js / JavaScript / TypeScript

- Function declarations: `function name() {}`
- Arrow functions: `const name = () => {}`
- Methods: `name() {}`
- Class methods: `async name() {}`
- Supports: `.js`, `.jsx`, `.ts`, `.tsx`, `.mjs` files

### Java

- Methods with various modifiers: `public static void method() {}`
- Supports: `.java` files

### Python

- Function definitions: `def function_name():`
- Methods: `def method_name(self):`
- Async functions: `async def function_name():`
- Type-annotated functions: `def function_name(arg: str) -> str:`
- Supports: `.py` files

## How It Works

1. **Repository Access**: Clones remote repositories to temporary directories or uses local paths
2. **Parallel Processing**: Scans multiple repositories concurrently for improved performance
3. **File Discovery**: Recursively finds all relevant source files (skips `node_modules`, `.git`, `target`, `build`, etc.)
4. **Test File Exclusion**: Automatically excludes test files to focus on production code
   - **Primary method**: Directory-based exclusion (most reliable)
     - Excludes files in `test`, `tests`, `__tests__`, `spec`, `specs` directories (case-insensitive)
     - Handles standard Java project structure (e.g., `src/test/java`)
     - Handles JavaScript test directories (e.g., `__tests__`)
   - **Secondary method**: Filename pattern matching (for edge cases)
     - Java: Files ending with `Test.java` or `Tests.java`
     - JavaScript/TypeScript: Files containing `.test.` or `.spec.`
     - Python: Files starting with `test_` or ending with `_test.py` or `_tests.py`
5. **Function Parsing**: Uses streaming parsers with regex patterns and brace/indentation tracking to identify function/method declarations
   - **Memory-Efficient**: Processes files line-by-line without loading entire files into memory, allowing analysis of very large files
   - **JavaScript/TypeScript/Java**: Counts lines by tracking brace pairs `{}`
   - **Python**: Counts lines by tracking indentation levels
6. **Size Calculation**: Counts lines from function start to end
7. **Filtering**: Applies minimum size filter to exclude trivial functions
8. **Ranking**: Sorts functions by line count and selects top N per repository
9. **Export**: Creates formatted Excel or JSON file with results and summary statistics
10. **Cleanup**: Automatically removes temporary cloned repositories

## Limitations

- Function size is measured by line count (including braces and blank lines)
- Nested functions are counted separately
- Very complex or unconventional syntax may not be detected
- Excludes common dependency directories (node_modules, target, build, etc.)
- Excludes test files based on common naming patterns

## Test Results

![Tests](https://github.com/Re4zOon/function-size-calculator/actions/workflows/test.yml/badge.svg)

### Test Summary

| Test Category | Tests | Status |
| -------------- | ------- | -------- |
| **Command-Line Arguments** | 1 | ✅ All Passed |
| **Excel Writer** | 6 | ✅ All Passed |
| **FunctionInfo** | 3 | ✅ All Passed |
| **JSON Writer** | 5 | ✅ All Passed |
| **Java Parser** | 4 | ✅ All Passed |
| **JavaScript/TypeScript Parser** | 5 | ✅ All Passed |
| **Python Parser** | 5 | ✅ All Passed |
| **Repository Scanner** | 3 | ✅ All Passed |
| **Test File Detection** | 13 | ✅ All Passed |
| **Test File Exclusion** | 4 | ✅ All Passed |
| **Total** | **54** | **✅ 54 Passed** |

### Performance

- **Execution Time**: ~0.21 seconds
- **Platform**: Linux, Python 3.12.3, pytest 9.0.2

*Last updated: 2026-01-05*

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
