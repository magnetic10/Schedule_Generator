@echo off
cd /d %~dp0
..\.venv\Scripts\python -m unittest discover -s tests
