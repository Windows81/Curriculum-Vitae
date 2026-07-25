# VisualPlugin's Cirriculum Vitæ

To render this résumé in PDF format, I use the nearest Chromium browser's headless mode.

You'll need:

- Python 3
  - The [get-chrome-paths module](https://github.com/Windows81/Get-Chrome-Paths)
- A Chromium-based browser that's easy to find
- A Bash-like shell to run [`_.sh`](_.sh)

```
pip3 install git+https://github.com/Windows81/Get-Chrome-Paths.git
```

## [`_.sh`](_.sh): How to Use

```
./_.sh gen
```

Paste a job description here, and receive a PDF file matching `./__cv_*.pdf` with its `<choice>` tags adjusted to reflect the word frequencies. Derived from the contents of `./cv/`.

```
./_.sh
```

Receive PDF files corresponding to the names and contents of relevant top-level directories (such as `./cv/` and `./cv-short/`).
