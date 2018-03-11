 def docx_replace_substring(doc_obj, template_str, value_str, max_cnt=100):
    was_found = True
    while (max_cnt > 0) and (was_found):
        was_found = False
        for p in doc_obj.paragraphs:
            pos = p.text.find(template_str)
            if pos >= 0:
                was_found = True
                inline = p.runs
                pos_end = len(template_str)
                # Loop added to work with runs (strings with same style)
                for i in range(len(inline)):
                    if pos_end == 0:
                        #replace completed
                        break
                    if (pos_end > 0) and (pos_end < len(template_str)):
                        #part of srting in another run, delete from start
                        part_len = min(len(inline[i].text), pos_end)  # length of part that in this run
                        inline[i].text = inline[i].text[part_len:]
                        pos_end = pos_end - part_len
                    if pos - len(inline[i].text) < 0:
                        # Use slicing to extract those parts of the original string to be kept
                        part_len = min(len(inline[i].text) - pos, len(template_str)) # length of part that in this run
                        inline[i].text = inline[i].text[:pos] + value_str + inline[i].text[(pos+part_len):]
                        pos_end = pos_end - part_len
                    else:
                        pos = pos - len(inline[i].text)
        max_cnt -= 1

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_substring(cell, template_str, value_str, max_cnt)
                
filename = 'doctestx.docx'
doc = docx.Document(filename)
docx_replace_substring(doc, r'text_to_replace', replace1)
docx_replace_substring(doc, r'text_to_replace2', replace2)
doc.save('generated_doc1.docx')                
