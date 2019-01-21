import xmltodict
from docx import Document
from docx.shared import Pt
from collections import OrderedDict

class EmilyCarrParser():
    def __init__(self, *args, **kwargs):
        self.data = xmltodict.parse(open('emily-carr.xml', 'rb'))

    def _parse(self):
        data = self.data['ead']['archdesc']['dsc']['c']
        data_list = []
        for item in data:
            output_data = {}
            if 'unittitle' in item['did'].keys():
                if isinstance(item['did']['unittitle'], list):
                    output_data['title'] = item['did']['unittitle'][0]['#text']
                    output_data['oti'] = item['did']['unittitle'][1]['#text']
                else:
                    output_data['title'] = item['did']['unittitle']['#text']
            if 'unitdate' in item['did'].keys():
                output_data['date'] = item['did']['unitdate']['#text']
            if 'note' in item['did'].keys():
                notes = item['did']['note']
                for note in notes:
                    if isinstance(note, str):
                        continue
                    if isinstance(note['p'], OrderedDict):
                        if '#text' not in note['p'].keys():
                            continue
                        if note['p']['#text'].startswith('Exhibited:'):
                            output_data['exhibit_note'] = note['p']['#text']
            if 'unitid' in item['did'].keys():
                output_data['reference_cd'] = item['did']['unitid']['#text']
            if 'physdesc' in item['did'].keys():
                output_data['physdesc'] = item['did']['physdesc']['#text']
            if 'scopecontent' in item.keys():
                if isinstance(item['scopecontent']['p'], OrderedDict):
                    output_data['scopecontent'] = item['scopecontent']['p']['#text']
                else:
                    output_data['scopecontent'] = item['scopecontent']['p']
            if 'acqinfo' in item.keys():
                output_data['acqinfo'] = item['acqinfo']['p']
            if 'relatedmaterial' in item.keys():
                if isinstance(item['relatedmaterial']['p'], OrderedDict):
                    output_data['related'] = item['relatedmaterial']['p']['#text']
                else:
                    output_data['related'] = item['relatedmaterial']['p']
            data_list.append(output_data)
        return data_list

    def write(self):
        data = self._parse()
        document = Document()
        for datum in data:
            text = 'Title: ' + datum['title']
            if 'date' in datum.keys():
                text += ' , Date: ' + datum['date'] + '\n '
            else:
                text += '\n '
            if 'oti' in datum.keys():
                text += 'Other Title Information: ' + datum['oti'] + ' \n '
            if 'reference_cd' in datum.keys():
                text += 'Identifier: ' + datum['reference_cd'] + ' \n '
            if 'physdesc' in datum.keys():
                text += 'Physical Description: ' + datum['physdesc'] + ' \n '
            if 'scopecontent' in datum.keys():
                text += 'Scope and Content: ' + datum['scopecontent'] + ' \n\n '
            if 'acqinfo' in datum.keys():
                text += 'Aquisition Information: ' + datum['acqinfo'] + ' \n '
            if 'related' in datum.keys():
                text += 'Related Items: ' + datum['related'] +' \n '
            if 'exhibit_note' in datum.keys():
                text += 'Exhibition Notes: ' + datum['exhibit_note']

            text += '\n'
            text+='___________________________________________________________________'


            paragraph = document.add_paragraph(text)
            run = paragraph.add_run()
            font = run.font
            font.name = 'Times New Roman'
            font.size = Pt(12)
        document.save('Emily Carr Art Collection.docx')
        return 0

if __name__ == '__main__':
    parse = EmilyCarrParser()
    parse.write()
