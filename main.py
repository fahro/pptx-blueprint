from  pptx_blueprint import Template

tpl = Template('data/example01.pptx')

shapes = tpl._find_shapes('*:title')

for shape in shapes:
    print(shape.text)