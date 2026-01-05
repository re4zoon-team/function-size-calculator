#!/usr/bin/env python3
"""
YAML to repos.txt converter.
Reads a YAML file containing repository information and extracts the URLs to repos.txt.
"""

import argparse
import sys

try:
    import yaml
except ImportError:
    print("Error: PyYAML is required. Install it with: pip install PyYAML")
    sys.exit(1)


def load_repos_from_yaml(yaml_file: str) -> list[str]:
    """
    Load repository URLs from a YAML file.

    Expected YAML structure:
    - name: spring-boot-template
      url: https://gitlab.local/services/spring-boot-template
      product: Devops
      type: springboot

    Args:
        yaml_file: Path to the YAML file

    Returns:
        List of repository URLs
    """
    with open(yaml_file, encoding='utf-8') as f:
        data = yaml.safe_load(f)

    if not isinstance(data, list):
        raise ValueError("YAML file must contain a list of repositories")

    urls = []
    for repo in data:
        if isinstance(repo, dict) and 'url' in repo:
            urls.append(repo['url'])

    return urls


def write_repos_to_file(urls: list[str], output_file: str) -> None:
    """
    Write repository URLs to a text file.

    Args:
        urls: List of repository URLs
        output_file: Path to the output file
    """
    with open(output_file, 'w', encoding='utf-8') as f:
        for url in urls:
            f.write(url + '\n')


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Extract repository URLs from a YAML file and write them to repos.txt'
    )
    parser.add_argument(
        'yaml_file',
        help='Path to the YAML file containing repository information'
    )
    parser.add_argument(
        '-o', '--output',
        default='repos.txt',
        help='Output file name (default: repos.txt)'
    )

    args = parser.parse_args()

    try:
        urls = load_repos_from_yaml(args.yaml_file)
        write_repos_to_file(urls, args.output)
        print(f"Successfully extracted {len(urls)} repository URLs to {args.output}")
    except FileNotFoundError:
        print(f"Error: File not found: {args.yaml_file}")
        sys.exit(1)
    except yaml.YAMLError as e:
        print(f"Error parsing YAML file: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
