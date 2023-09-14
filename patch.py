from datetime import datetime
import win32com.client
import zipfile
import os
import re

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

WORK_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="work-title">
			{work_title}<text:span text:style-name="work-sub">{work_sub}</text:span>
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


def make_projects(dir_path: str, file: str) -> list[str]:
    path = f'{dir_path}/_projects.txt'
    with open(path, 'r', encoding='utf-8') as f:
        data_lines = f.readlines() + ['\n']
        data = list[list[str]]()
        prev_i = 0
        for i, l in enumerate(data_lines):
            if l != '\n':
                continue
            data.append([
                l.rstrip('\n')
                for l in data_lines[prev_i:i]
            ])
            prev_i = i + 1

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


def make_work_exp(dir_path: str, file: str) -> list[str]:
    path = f'{dir_path}/_work.txt'
    with open(path, 'r', encoding='utf-8') as f:
        data_lines = f.readlines() + ['\n']
        data = list[list[str]]()
        prev_i = 0
        for i, l in enumerate(data_lines):
            if l != '\n':
                continue
            data.append([
                l.rstrip('\n')
                for l in data_lines[prev_i:i]
            ])
            prev_i = i + 1

    project_fill = '\n'.join(
        WORK_TEMPLATE.format_map({
            'work_sub': d[2],
            'work_title': d[1],
            'work_date': d[0],
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


def update_date(dir_path: str, file: str) -> list[str]:
    date = datetime.utcnow().strftime('%Y-%m-%d %HZ')
    return [
        f'<text:p text:style-name="foot-date">{date}</text:p>',
    ]


def xml_mod(dir_path: str, file: str, methods: dict[str, callable]):
    path = f'{dir_path}/{file}'
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    ranges = {}
    for i, l in enumerate(lines):

        # Captures XML comments.
        m = re.match('^(\t*)<!-- (.+) (.+) -->\s*$', l)
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

    # Replace text between the delimiting lines with what we return from calls to functions within METHODS.
    for (i, d) in reversed(ranges.items()):
        lines[d['start']:d['end']] = [
            f"{d['tab_prefix']}{l}\n"
            for l in methods[i](dir_path, file)
        ]

    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(lines)


def dir_mod(dir_path: str):
    xml_mod(dir_path, 'content.xml', {
        'PROJECTS': make_projects,
        'WORK': make_work_exp,
    })
    xml_mod(dir_path, 'styles.xml', {
        'UPDATE-DATE': update_date,
    })


def dir2odt(dir_path: str, odt_path: str):
    with zipfile.ZipFile(odt_path, 'w') as archive:
        for f in os.listdir(dir_path):
            archive.write(f'{dir_path}/{f}', f)


def odt2pdf(odt_path: str, pdf_path: str):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(os.path.abspath(odt_path))

    # 17 is an alias for PDF.
    doc.SaveAs(os.path.abspath(pdf_path), 17)
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    dir_mod('src')
    dir2odt('src', 'cv.odt')
    odt2pdf('cv.odt', 'cv.pdf')
