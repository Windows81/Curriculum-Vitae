#!/bin/bash
temp=_current_cv_profile.txt
dir="$(realpath -e "$(dirname "$0")")"
print_cv() {
    echo -n "$1" > "$temp"
    python3 - "$dir" "$dir/$1.pdf" "$1" <<'PY'
import sys
from playwright.sync_api import sync_playwright

base, out, profile = sys.argv[1], sys.argv[2], sys.argv[3]
with sync_playwright() as p:
    browser = p.chromium.launch(args=["--allow-file-access-from-files"])
    page = browser.new_page()
    page.goto("file://" + base + "/index.html?cvProfile=" + profile)
    page.wait_for_timeout(800)
    page.pdf(path=out, prefer_css_page_size=True, print_background=True)
    browser.close()
PY
    rm "$temp"
}

case "$1" in
    gen)
        cv_name="__cv_$(uuidgen)"
        read -r -d '`' \
        -p 'Enter job description; stops at "`": ' \
        job_desc
        cp -r "cv" "$cv_name"
        echo "$job_desc" > "$cv_name/_job_desc.txt"
        print_cv "$cv_name"
        rm -r -f "$cv_name"
        ;;
    all)
        for cv_name in $(
            find.exe . \
                -maxdepth 1 -type d ! \
                -name old ! -name '.*' \
                -printf "%f\n"
            ); do
            print_cv "$cv_name"
        done
        ;;
    "")
        print_cv "cv"
        ;;
    *)
        print_cv "$1"
        ;;
esac