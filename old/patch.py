<<<<<<< HEAD:patch.py
from ast import arg
from dataclasses import dataclass
import functools
=======
>>>>>>> 03ac26b (2026-02-02T0542Z):old/patch.py
from typing import Callable
import win32com.client
import datetime
import zipfile
import os
import re

SOURCE_DIR = '_src'


@functools.cache
def resolve_text_file(contents_dir: str, file_name: str) -> str:
    while True:
        path = os.path.join(contents_dir, file_name)
        with open(path, 'r', encoding='utf-8') as f:
            text = f.readline().rstrip('\r\n')
            if not text.startswith('@'):
                return path
            contents_dir = os.path.relpath(text[1:])


@functools.cache
def load_template(source_dir: str, file_name: str) -> str:
    with open(os.path.join(source_dir, '_templates', file_name)) as f:
        return f.read()


@functools.cache
def load_xml_lines(contents_dir: str, file_name: str) -> list[list[str]]:
    path = resolve_text_file(contents_dir, file_name)
    with open(path, 'r', encoding='utf-8') as f:
        data_lines = f.readlines() + ['\n']
        data = list[list[str]]()
        prev_i = 0
        for i, l in enumerate(data_lines):
            if l != '\n':
                continue
            data.append([
                l.rstrip('\r\n').lstrip('- ')
                for l in data_lines[prev_i:i]
                if not l.startswith('#')
            ])
            prev_i = i + 1
    return data


def make_skills(contents_dir: str, source_dir: str) -> list[str]:
    template = load_template(source_dir, 'skills.txt')
    data = load_xml_lines(contents_dir, 'skills.txt')

    fill = '\n'.join(
        template.format_map({
            'tech_skill': l,
        })
        for d in data for l in d
    )
    return fill.split('\n')


<<<<<<< HEAD:patch.py
def make_education(contents_dir: str, source_dir: str) -> list[str]:
    template = load_template(source_dir, 'education.txt')
    template_desc = load_template(source_dir, 'education-desc.txt')
    data = load_xml_lines(contents_dir, 'education.txt')
=======
EDUCATION_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="edu-title">{educ_title}
		</text:p>
		<text:list
			text:style-name="edu-list"
			text:continue-numbering="true">
{educ_descs}
		</text:list>
	</table:table-cell>
	<table:table-cell table:style-name="date-cell">
		<text:p text:style-name="date">{educ_date}</text:p>
	</table:table-cell>
</table:table-row>
""".lstrip('\n')

EDUCATION_TEMPLATE_DESC = """
			<text:list-item>
				<text:p text:style-name="edu-desc">{educ_desc}</text:p>
			</text:list-item>
""".lstrip('\n')


def make_education(contents_dir: str) -> list[str]:
    data = process_lines(resolveTextFile(contents_dir, 'education.txt'))
>>>>>>> 03ac26b (2026-02-02T0542Z):old/patch.py

    fill = '\n'.join(
        template.format_map({
            'educ_date': d[0],
            'educ_title': d[1],
            'educ_descs': ''.join(
                template_desc.format_map({
                    'educ_desc': desc
                })
                for desc in d[2:]
            ),
        })
        for d in data
    )
    return fill.split('\n')


def make_projects(contents_dir: str, source_dir: str) -> list[str]:
    template = load_template(source_dir, 'projects.txt')
    template_desc = load_template(source_dir, 'projects-desc.txt')
    data = load_xml_lines(contents_dir, 'projects.txt')

    fill = '\n'.join(
        template.format_map({
            'proj_sub': d[3],
            'proj_link': d[2],
            'proj_title': d[1],
            'proj_date': d[0],
            'proj_descs': ''.join(
                template_desc.format_map({
                    'proj_desc': desc
                })
                for desc in d[4:]
            ),
        })
        for d in data
    )
    return fill.split('\n')


def make_work_exp(contents_dir: str, source_dir: str) -> list[str]:
    template = load_template(source_dir, 'work.txt')
    template_desc = load_template(source_dir, 'work-desc.txt')
    data = load_xml_lines(contents_dir, 'work.txt')

    project_fill = '\n'.join(
        template.format_map({
            'work_date': d[0],
            'work_title': d[1],
            'work_sub': d[2],
            'work_descs': ''.join(
                template_desc.format_map({
                    'work_desc': desc
                })
                for desc in d[3:]
            ),
        })
        for d in data
    )
    return project_fill.split('\n')


def update_date(contents_dir: str, source_dir: str) -> list[str]:
    date = datetime.datetime.fromtimestamp(
        max(
            os.path.getmtime(path)
            for fn in os.listdir(contents_dir)
            if (path := os.path.join(contents_dir, fn))
            and path.endswith('.txt')
        ),
        tz=datetime.UTC,
    )
    date_str = date.strftime('%Y-%m-%d %HZ')
    return [
        f'<text:p text:style-name="foot-date">{date_str}</text:p>',
    ]


def modify_xml(
    contents_dir: str,
    source_dir: str,
    file_name: str,
    methods: dict[str, Callable[[str, str], list[str]]]
) -> None:
    @dataclass
    class Range:
        start: int
        end: int
        tab_prefix: str

    xml_path = os.path.join(source_dir, file_name)
    with open(xml_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    ranges: dict[str, Range] = {}
    for i, l in enumerate(lines):

        # Captures XML comments.
        m = re.match(r'^(\t*)<!-- (.+) (.+) -->\s*$', l)
        if not m:
            continue
        typ = m.group(2)
        mode = m.group(3)

        if mode == 'START':
            ranges[typ] = Range(
                start=i + 1,
                end=i + 1,
                tab_prefix=m.group(1),
            )

        elif mode == 'END':
            ranges[typ].end = i

    # Replaces text between the delimiting lines with what we return from calls to functions within METHODS.
    for (i, d) in reversed(ranges.items()):
        lines[d.start:d.end] = [
            f"{d.tab_prefix}{l}\n"
            for l in methods[i](contents_dir, source_dir)
        ]

    with open(xml_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def modify_dir(contents_dir: str, source_dir: str) -> None:
    modify_xml(contents_dir, source_dir, 'content.xml', {
        'SKILLS': make_skills,
        'EDUCATION': make_education,
        'PROJECTS': make_projects,
        'WORK': make_work_exp,
    })
    modify_xml(contents_dir, source_dir, 'styles.xml', {
        'UPDATE-DATE': update_date,
    })


def dir_to_odt(contents_dir: str, odt_path: str) -> None:
    with zipfile.ZipFile(odt_path, 'w') as archive:
        for f in os.listdir(contents_dir):
            archive.write(os.path.join(contents_dir, f), f)


def odt_to_pdf(odt_path: str, pdf_path: str):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(odt_path)

    # 17 is an alias for PDF.
    doc.SaveAs(pdf_path, 17)
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    module_dir_abs = os.path.abspath(os.path.dirname(__file__))

    import argparse
    parser = argparse.ArgumentParser()
    _ = parser.add_argument(
        'contents_dirs',
        default=[
            f.name
            for f in os.scandir(module_dir_abs)
            if f.name != SOURCE_DIR
            and f.is_dir(follow_symlinks=False)
            and not f.name.startswith(('.', '_'))
        ],
        nargs='*',
        type=str,
    )

    args = parser.parse_args()
    contents_dirs: list[str] = args.contents_dirs

    source_names = [
        f.name
        for f in os.scandir(os.path.join(module_dir_abs, SOURCE_DIR))
        if f.is_dir(follow_symlinks=False)
        and not f.name.startswith('.')
    ]

    def f(contents_dir: str, source_name: str) -> None:
        contents_dir_abs = os.path.abspath(contents_dir)
        source_dir = os.path.join(module_dir_abs, SOURCE_DIR, source_name)
        odt_path = os.path.join(contents_dir_abs, f'{source_name}.odt')
        pdf_path = os.path.join(contents_dir_abs, f'{source_name}.pdf')
        modify_dir(contents_dir, source_dir)
        dir_to_odt(source_dir, odt_path)
        odt_to_pdf(odt_path, pdf_path)
        print(f'Finished "{contents_dir}" ({source_name})')

    for c in contents_dirs:
        for s in source_names:
            f(c, s)
