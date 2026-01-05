#!/usr/bin/env python3
"""Pytest suite for yaml_to_repos.py."""

import os
import sys
from pathlib import Path

import pytest

# Add parent directory to path to import the module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from yaml_to_repos import load_repos_from_yaml, write_repos_to_file


class TestLoadReposFromYaml:
    """Tests for load_repos_from_yaml function."""

    def test_load_valid_yaml(self, tmp_path: Path):
        """Should load valid YAML file with repository URLs."""
        yaml_content = """
- name: spring-boot-template
  url: https://gitlab.local/services/spring-boot-template
  product: Devops
  type: springboot
- name: spring-boot
  url: https://gitlab.local/services/spring-boot
  product: Devops
  type: springboot
"""
        yaml_file = tmp_path / "repos.yaml"
        yaml_file.write_text(yaml_content)

        urls = load_repos_from_yaml(str(yaml_file))

        assert len(urls) == 2
        assert urls[0] == "https://gitlab.local/services/spring-boot-template"
        assert urls[1] == "https://gitlab.local/services/spring-boot"

    def test_load_yaml_with_missing_url(self, tmp_path: Path):
        """Should skip entries without URL field."""
        yaml_content = """
- name: repo-with-url
  url: https://gitlab.local/services/repo-with-url
  product: Devops
- name: repo-without-url
  product: Devops
"""
        yaml_file = tmp_path / "repos.yaml"
        yaml_file.write_text(yaml_content)

        urls = load_repos_from_yaml(str(yaml_file))

        assert len(urls) == 1
        assert urls[0] == "https://gitlab.local/services/repo-with-url"

    def test_load_empty_yaml(self, tmp_path: Path):
        """Should handle empty YAML file."""
        yaml_file = tmp_path / "repos.yaml"
        yaml_file.write_text("[]")

        urls = load_repos_from_yaml(str(yaml_file))

        assert len(urls) == 0

    def test_load_yaml_file_not_found(self):
        """Should raise FileNotFoundError for non-existent file."""
        with pytest.raises(FileNotFoundError):
            load_repos_from_yaml("/nonexistent/repos.yaml")

    def test_load_yaml_invalid_structure(self, tmp_path: Path):
        """Should raise ValueError for non-list YAML structure."""
        yaml_content = """
name: single-repo
url: https://gitlab.local/services/single-repo
"""
        yaml_file = tmp_path / "repos.yaml"
        yaml_file.write_text(yaml_content)

        with pytest.raises(ValueError, match="must contain a list"):
            load_repos_from_yaml(str(yaml_file))


class TestWriteReposToFile:
    """Tests for write_repos_to_file function."""

    def test_write_repos_to_file(self, tmp_path: Path):
        """Should write URLs to file."""
        urls = [
            "https://gitlab.local/services/repo1",
            "https://gitlab.local/services/repo2",
        ]
        output_file = tmp_path / "repos.txt"

        write_repos_to_file(urls, str(output_file))

        content = output_file.read_text()
        lines = content.strip().split("\n")

        assert len(lines) == 2
        assert lines[0] == "https://gitlab.local/services/repo1"
        assert lines[1] == "https://gitlab.local/services/repo2"

    def test_write_empty_list(self, tmp_path: Path):
        """Should create empty file for empty list."""
        output_file = tmp_path / "repos.txt"

        write_repos_to_file([], str(output_file))

        content = output_file.read_text()
        assert content == ""
