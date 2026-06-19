#!/bin/bash
temp=_current_cv_profile.txt
dir="$(realpath -e "$(dirname "$0")")"
exe="$(python3 -m get_chrome_paths | head -n 1)"

print_cv() {
    echo -n "$1" > "$temp"
    $exe --headless --allow-chrome-scheme-url \
    --print-to-pdf="$dir/$1.pdf" --disable-web-security \
    --allow-file-access-from-files --timeout=5000 \
    --homepage "$dir/index.html"
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
    *)
        for cv_name in $(
            find.exe . \
                -maxdepth 1 -type d ! \
                -name old ! -name '.*' \
                -printf "%f\n"
            ); do
            print_cv "$cv_name"
        done
        ;;
esac