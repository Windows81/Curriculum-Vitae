import re
PROJECT_TEMPLATE = """
<table:table-row table:style-name="sect-row">
	<table:table-cell table:style-name="sect-cell">
		<text:p text:style-name="proj-head">
			<text:span text:style-name="proj-title">{proj_title}</text:span> {proj_sub}</text:p>
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

if __name__ == '__main__':
    path = 'src/content.xml'
    path2 = 'src/projects.txt'
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    tab_prefix = ''
    start_i = end_i = 0
    for i, l in enumerate(lines):
        m = re.match('^(\t+)<!-- PROJECTS (.+) -->\s+$', l)
        if not m:
            continue

        mode = m.group(2)
        if mode == 'START':
            tab_prefix = m.group(1)
            start_i = i + 1

        elif mode == 'END':
            end_i = i
            break

    with open(path2, 'r', encoding='utf-8') as f:
        project_data_lines = f.readlines()
        project_data = []
        prev_i = 0
        for i, l in enumerate(project_data_lines):
            if l != '\n':
                continue
            project_data.append(project_data_lines[prev_i:i])
            prev_i = i + 1

    project_fill = '\n'.join(
        PROJECT_TEMPLATE.format_map({
            'proj_title': proj[1].rstrip('\n'),
            'proj_sub': proj[2].rstrip('\n'),
            'proj_date': proj[0].rstrip('\n'),
            'proj_descs': '\n'.join(
                PROJECT_TEMPLATE_DESC.format_map({
                    'proj_desc': desc.rstrip('\n')
                })
                for desc in proj[3:]
            ),
        })
        for proj in project_data
    )
    project_lines = [f'{tab_prefix}{l}\n' for l in project_fill.split('\n')]

    new_lines = lines[:start_i] + project_lines + lines[end_i:]
    with open(path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
