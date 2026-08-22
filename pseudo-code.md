def login(str: my_login, str: my_pwd):
    open_page = navigate(“https://admin.reading-space.com/login”)
    user_field = open_page.locate(element_name=<input>,id=”username”)
    text_input(field=user_field, data=my_login)
    pwd_field= open_page.locate(element_name=<input>,id=”password”)
    text_input(field=pwd_field, data=my_pwd)
    submit_button = open_page.locate(element_name=<button>, type=”submit”)
    submit_button.click()


def batch_process(function: record_fn, path: spreadsheet_path):
    id_col = locate_header(“id”)
    for id in id_col:
        record_fn(id)   

def oec_record(str: book_id, path: spreadsheet_path):
    with open(spreadsheet_path) as f:
        f.active_sheet = “OEC”
        id_col = locate_header(“id”)
        row_n = locate_row(search_key=book_id, lookup_range=id_col:id_col)
        open_page = navigate(f”https://admin.reading-space.com/activities/{book_id}/open-ended-questions/edit”)
            For elem in open_page.elements(): 
                if elem.name == “<textarea>” and elem.placeholder == ”What do you think this story will be about?”
                    el_data = elem.text_content
                    q_col = f.locate_header(“Q1”)
                    for i in range(10):
                        if f.cell[row_n][q_col] == none:
                            q_col = q_col+i
                            break
                    f.insert(row=row_n, column=q_col, data=el_data) 

def cc_record(str: book_id, path: spreadsheet_path):
    with open(spreadsheet_path) as f:
        f.active_sheet = “CC”
        id_col = f.locate_header(“id”)
        row_n = f.locate_row(search_key=book_id, lookup_range=id_col:id_col)
        open_page = navigate(f”https://admin.reading-space.com/activities/{book_id}/context-clue/edit”)
            For elem in open_page.elements(): 
                if elem.name == “<textarea>”:
                    el_data = elem.text_content
                    q_col = None
                    for i in range(1, 11):
                        q_col = f.locate_header(f“Q{i}”)
                        if f.cell[row_n][q_col] == None:
                            f.insert(row=row_n, column=q_col, data=el_data) 
                            break
            For elem in open_page.elements(): 
                if elem.class == “flex items-center gap-2 rounded-full bg-neutral-100 px-3 py-1 text-sm text-gray-900”:
                    text_span = elem.find_element(“span”)
                    el_data = text_span.text_content
                    a_col = None
                    for i in range(1, 11):
                        a_col = f.locate_header(f“A{i}”)
                        if f.cell[row_n][a_col] == None:
                            f.insert(row=row_n, column=a_col, data=el_data) 
                            break

def wac_record(str: book_id, path: spreadsheet_path)
    with open(spreadsheet_path) as f:
        f.active_sheet = “CC”
        id_col = f.locate_header(“id”)
        row_n = f.locate_row(search_key=book_id, lookup_range=id_col:id_col)
        open_page = navigate(f”https://admin.reading-space.com/activities/{book_id}/context-clue/edit”)
        tables = [el for el in open_page.elements() if el.name == “table”][0]
        for tr in table:
            vocab_word = tr[1].text_content ### meaning the second td in tr
            vocab_pos = tr[2].text_content ### meaning the third td in tr
            vocab_def = tr[3].text_content ### meaning the fourth td in tr
            for i in range(1, 11):
                w_col = f.locate_header(f“W{i}”)
                pos_col = f.locate_header(f“POS{i}”)
                def_col = f.locate_header(f“DEF{i}”)
                if f.cell[row_n][w_col] == None:
                    f.insert(row=row_n, column=a_col, data=vocab_word) 
                    f.insert(row=row_n, column=pos_col, data=vocab_pos)
                    f.insert(row=row_n, column=def_col, data=vocab_def 

def tmc_record(str: book_id, path: spreadsheet_path)
    with spreadsheet_path open as f:
        f.active_sheet = “TMC”
        id_col = f.locate_header(“id”)
        row_n = f.locate_row(search_key=book_id, lookup_range=id_col:id_col)
        open_page = navigate(f”https://admin.reading-space.com/activities/{book_id}/text-multiple-choice/edit”)
        tables = [el for el in open_page.elements() if el.class == “flex flex-col gap-4 rounded-lg bg-white p-6 shadow-md”]
        for table in tables:
            question_elem = None
            answers_wrap = None
            answers_elem = None
            q_text = None
            for elem in table.elements():
                If elem.class == “flex flex-col gap-2” and elem.name = “label”:
                    question_elem = elem
                If elem.class == “flex flex-col gap-2” and elem.name = “div”:
                    answers_wrap = elem
            For elem in question_elem.elements:
                If elem.name = “textarea”:
                    q_text = elem.text_content
            For elem in answers_wrap.elements:
                If elem.class == “flex flex-col gap-3”:
                    answers_elem = elem
            idv_el_list = answers_elem.elements
            txt_ans_list = []
            For i in range(4):
                idv_ans = idv_el_list[i]:
                for elem in idv_ans.elements:
                    if elem.name == “input”:
                        txt_ans_list.append(elem.value)
            ans_a = txt_ans_list[0]
            ans_b = txt_ans_list[1]
            ans_c = txt_ans_list[2]
            ans_d = txt_ans_list[3]
            for i in range(1, 11):
                q_col = f.locate_header(f“Q{i}”)
                ans_a_col= f.locate_header(f“AnsA{i}”)
                ans_b_col= f.locate_header(f“AnsB{i}”)
                ans_c_col= f.locate_header(f“AnsC{i}”)
                ans_d_col= f.locate_header(f“AnsD{i}”)
                if f.cell[row_n][q_col] == None:
                    f.insert(row=row_n, column=q_col, data=q_text)
                    f.insert(row=row_n, column=ans_a_col, data=ans_a) 
                    f.insert(row=row_n, column=ans_b_col, data=ans_b)
                    f.insert(row=row_n, column=ans_c_col, data=ans_c)
                    f.insert(row=row_n, column=ans_d_col, data=ans_d)
