import win32com.client
import datetime
import zipfile
import os
import re


def resolveTextFile(contents_dir: str, file_name: str) -> str:
    while True:
        path = f'{contents_dir}/{file_name}'
        with open(path, 'r', encoding='utf-8') as f:
            text = f.readline().rstrip('\r\n')
            if not text.startswith('@'):
                return path
            contents_dir = text[1:]


def process_lines(path: str) -> list[list[str]]:
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


SKILL_TEMPLATE = """
<text:list-item>
	<text:p text:style-name="tech-skill">{tech_skill}</text:p>
</text:list-item>
""".lstrip('\n')


def make_skills(contents_dir: str) -> list[str]:
    data = process_lines(resolveTextFile(contents_dir, 'skills.txt'))

    fill = '\n'.join(
        SKILL_TEMPLATE.format_map({
            'tech_skill': l,
        })
        for d in data for l in d
    )
    return fill.split('\n')


EDUCATION_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="edu-title">{educ_title}<text:span text:style-name="edu-sub"> – {educ_sub}</text:span>
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

    fill = '\n'.join(
        EDUCATION_TEMPLATE.format_map({
            'educ_date': d[0],
            'educ_title': d[1],
            'educ_sub': d[2],
            'educ_descs': ''.join(
                EDUCATION_TEMPLATE_DESC.format_map({
                    'educ_desc': desc
                })
                for desc in d[3:]
            ),
        })
        for d in data
    )
    return fill.split('\n')


PROJECT_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="proj-head">
			<text:a xlink:href="{proj_link}">
				<text:span text:style-name="proj-title">
				{proj_title}</text:span>
			</text:a>{proj_sub}</text:p>
		<text:list
			text:style-name="proj-list"
			text:continue-numbering="true">
{proj_descs}
		</text:list>
	</table:table-cell>
	<table:table-cell table:style-name="date-cell">
		<text:p text:style-name="date">{proj_date}</text:p>
	</table:table-cell>
</table:table-row>
""".lstrip('\n')

PROJECT_TEMPLATE_DESC = """
			<text:list-item>
				<text:p text:style-name="proj-desc">{proj_desc}</text:p>
			</text:list-item>
""".lstrip('\n')


def make_projects(contents_dir: str) -> list[str]:
    data = process_lines(resolveTextFile(contents_dir, 'projects.txt'))

    fill = '\n'.join(
        PROJECT_TEMPLATE.format_map({
            'proj_sub': d[3],
            'proj_link': d[2],
            'proj_title': d[1],
            'proj_date': d[0],
            'proj_descs': ''.join(
                PROJECT_TEMPLATE_DESC.format_map({
                    'proj_desc': desc
                })
                for desc in d[4:]
            ),
        })
        for d in data
    )
    return fill.split('\n')


WORK_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="work-title">
			{work_title}<text:span text:style-name="work-sub"> {work_sub}</text:span>
		</text:p>
{work_descs}
	</table:table-cell>
	<table:table-cell table:style-name="date-cell">
		<text:p text:style-name="date">{work_date}</text:p>
	</table:table-cell>
</table:table-row>
""".lstrip('\n')

WORK_TEMPLATE_DESC = """
		<text:p text:style-name="work-desc">
			{work_desc}
		</text:p>
""".lstrip('\n')


def make_work_exp(contents_dir: str) -> list[str]:
    data = process_lines(resolveTextFile(contents_dir, 'work.txt'))

    project_fill = '\n'.join(
        WORK_TEMPLATE.format_map({
            'work_date': d[0],
            'work_title': d[1],
            'work_sub': d[2],
            'work_descs': ''.join(
                WORK_TEMPLATE_DESC.format_map({
                    'work_desc': desc
                })
                for desc in d[3:]
            ),
        })
        for d in data
    )
    return project_fill.split('\n')


def update_date(contents_dir: str) -> list[str]:
    date = datetime.datetime.now(datetime.UTC).strftime('%Y-%m-%d %HZ')
    return [
        f'<text:p text:style-name="foot-date">{date}</text:p>',
    ]


def xml_mod(contents_dir: str, source_dir: str, file_name: str, methods: dict[str, callable]):
    xml_path = f'{source_dir}/{file_name}'
    with open(xml_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    ranges = {}
    for i, l in enumerate(lines):

        # Captures XML comments.
        m = re.match(r'^(\t*)<!-- (.+) (.+) -->\s*$', l)
        if not m:
            continue
        typ = m.group(2)
        mode = m.group(3)

        if mode == 'START':
            ranges[typ] = {
                'start': i + 1,
                'end': i + 1,
                'tab_prefix': m.group(1),
            }

        elif mode == 'END':
            ranges[typ]['end'] = i

    # Replaces text between the delimiting lines with what we return from calls to functions within METHODS.
    for (i, d) in reversed(ranges.items()):
        lines[d['start']:d['end']] = [
            f"{d['tab_prefix']}{l}\n"
            for l in methods[i](contents_dir)
        ]

    with open(xml_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def dir_mod(contents_dir: str, source_dir: str):
    xml_mod(contents_dir, source_dir, 'content.xml', {
        'SKILLS': make_skills,
        'EDUCATION': make_education,
        'PROJECTS': make_projects,
        'WORK': make_work_exp,
    })
    xml_mod(contents_dir, source_dir, 'styles.xml', {
        'UPDATE-DATE': update_date,
    })


def dir2odt(contents_dir: str, odt_path: str):
    with zipfile.ZipFile(odt_path, 'w') as archive:
        for f in os.listdir(contents_dir):
            archive.write(f'{contents_dir}/{f}', f)


def odt2pdf(odt_path: str, pdf_path: str):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(os.path.abspath(odt_path))

    # 17 is an alias for PDF.
    doc.SaveAs(os.path.abspath(pdf_path), 17)
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'contents_dir',
        default='default',
        nargs='?',
        type=str,
    )
    args = parser.parse_args()
    contents_dir = args.contents_dir
    source_dir = f'{os.path.abspath(os.path.dirname(__file__))}/src'

    dir_mod(contents_dir, source_dir)
    dir2odt(source_dir, f'{contents_dir}/cv.odt')
    odt2pdf(f'{contents_dir}/cv.odt', f'{contents_dir}/cv.pdf')
