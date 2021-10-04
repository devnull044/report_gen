import validators, dateparser, yaml
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from docx.shared import Cm
from docxtpl import DocxTemplate
import sys

def gen_report(file):
    """ Load the data from the YAML file """
    data = yaml.safe_load(open(file))
    #data = data.replace("\n","<br>")
    print(data)
    """ Get number of findings """
    stats = [{'critical':0,
            'high':0,
            'medium':0,
            'low':0}]
    if data['findings']:
        for x in data['findings']:
            if "critical" in x.keys():
                stats[0]['critical'] += 1
            elif "high" in x.keys():
                stats[0]['high'] += 1
            elif "medium" in x.keys():
                stats[0]['medium'] += 1
            elif "low" in x.keys():
                stats[0]['low'] += 1

    data['stats'] = stats[0]
    print(data)
    """ Retrieve the templates """
    #YAML piece
    env = Environment(loader = FileSystemLoader('.'), trim_blocks=True, lstrip_blocks=True)
    template = env.get_template("template.html")

    #Word piece
    word_template = DocxTemplate("template.docx")
    word_template.render(data)
    word_template.save("final.docx")

    # Fills in the variables in the HTML file
    template.stream(data).dump("pre_test.html")
    HTML("pre_test.html").write_pdf('test.pdf')

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage [+]: %s <file.yml>' % sys.argv[0])
        exit()

    file = sys.argv[1]
    gen_report(file)
