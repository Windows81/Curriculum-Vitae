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
./_.sh $NAME
```

Paste a directory _name_ (such as `cv` or `cv-short`) and receive a PDF file whose fields are filled from that directory's files.

---

```
./_.sh
```

Receive a PDF file whose fields are filled from the `./cv/` directory's files.

---

```
./_.sh all
```

Receive PDF files whose fields are filled from _every_ directory's files (e.g. `./cv/`, `./cv-short/`).

---

```
./_.sh gen
```

Paste a job description to _stdin_, and receive a PDF file matching `./__cv_*.pdf` with its `<choice>` tags adjusted to reflect the word frequencies. Derived from the contents of `./cv/`.
