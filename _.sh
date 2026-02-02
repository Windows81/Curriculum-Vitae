#!/bin/sh
dir="$(realpath -e "$(dirname "$0")")"
exe="$(python3 -m get_chrome_paths | head -n 1)"

print_cv() {
    echo -n "$1" > _current_cv_profile.txt
    $exe --headless --allow-chrome-scheme-url --print-to-pdf="$dir/test.pdf" --disable-web-security --allow-file-access-from-files --timeout=5000 --homepage "$dir/index.html"
}

for f in $(
    find.exe . -maxdepth 1 -type d ! -name old ! -name '.*' -printf "%f\n"
    ); do
    print_cv "$f"
done