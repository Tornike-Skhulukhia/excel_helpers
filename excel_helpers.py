'''
    ფუნქციები Excel-ის(xlsx) ფაილებთან მუშაობის გასამარტივებლად
'''


def cell_is_empty(cell_value):
    '''
    ამოწმებს, არის თუ არა Excel-ის ფაილის უჯრა ცარიელი
    არგუმენტები:
        1. cell_value - xlsx ფაილში უჯრის მნიშვნელობა
    '''
    return str(cell_value).strip() in ['None', 'NONE', "", None]


def get_sheet_names(file, return_wb=False):
    '''
    გვიბრუნებს ფურცლების სიას კონკრეტული დოკუმენტიდან.
    თუ return_wb True-ა, დაბრუნებული შედეგის მე-2
    ელემენტი გახსნილი workbook ობიექტია.

    არგუმენტები:
        1. file - xlsx ფაილის მისამართი
        2. return_wb - დაგვიბრუნოს თუ არა workbook ობიექტი
                (ნაგულისხმევად=False). გამოსადეგია
                სისწრაფის თვალსაზრისით დიდი დოკუმენტებისთვის.
    '''
    import openpyxl

    wb = openpyxl.load_workbook(file)

    return wb.sheetnames if not return_wb else [wb.sheetnames, wb]


def get_last_row_num(file_or_wb_obj, sheet_name, column,
                     number=10, start_row=1):
    '''
        Excel-ის(xlsx) ფაილიდან ან შესაბამისი ობიექტიდან
    (openpyxl.workbook.workbook.Workbook)
    გვიბრუნებს ჩვენთვის სასურველი ფურცლის(sheet) სასურველი სვეტის(column)
    ბოლო სტრიქონის ნომერს.

        ბოლოდ მიიჩნევა ის სტრიქონი, რომლის ქვემოთაც გვხვდება მინიმუმ
    განსაზღვრული რაოდენობის(number) ცარიელი უჯრები.

    არგუმენტები:
        1. file_or_wb_obj  -  xlsx ფაილის მისამართი, ან ფაილის ობიექტი
                            (openpyxl.workbook.workbook.Workbook)

        2. sheet_name      -  ფურცლის(sheet) სახელი

        3. column          -  რომელი სვეტის გამოყენება გვინდა
                            ბოლო სტრიქონის საპოვნელად

        4. number          -  (ნაგულისხმევად = 10),
                            რამდენი ქვედა ცარიელი უჯრა ჩავთვალოთ
                            საკმარისად, რათა უჯრა ბოლოდ მივიჩნიოთ

        5. start_row       -  (ნაგულისხმევად = 1),
                            საწყისი სტრიქონი(შესაძლოა მონაცემები არ
                            იწყებოდეს პირველივე სტრიქონიდან)
    '''
    import openpyxl
    # load
    if not isinstance(file_or_wb_obj, openpyxl.workbook.workbook.Workbook):
        wb = openpyxl.load_workbook(file_or_wb_obj)
    else:
        wb = file_or_wb_obj

    # worksheet
    ws = wb[sheet_name]

    # to make code shorter
    r = start_row
    c = column

    # empty_num = 0
    empties_num = sum(
        [cell_is_empty(ws[f'{c}{r + i}'].value)
         for i in range(1, number + 1)])
    # if first is empty and there in next rows also
    # no value found, return 0
    if empties_num == number and cell_is_empty(ws[f'{c}{r}'].value):
        return 0
    else:
        # count better
        while empties_num != number:
            r += 1
            empties_num = sum(
                    [cell_is_empty(ws[f'{c}{r + i}'].value)
                     for i in range(1, number + 1)])
    return r


def get_data(
            file_or_wb_obj,
            check_column,
            start_row,
            data_columns,
            sheet_index=0,
            data_only=False,
            number=10):
    '''
        ფუნქცია გვეხმარება ჩვენთვის სასურველი სვეტების მონაცემების
    მიღებაში Excel-ის(xlsx) ფაილებიდან.

        ჩვენ ვუთითებთ საწყის სტრიქონს და სვეტებს, რომლებიდანაც
    გვინდა მონაცემების მიღება.

        მონაცემები ბრუნდება list ტიპად, ქვე-list-ებით თითოეული
    სტრიქონისთვის, სვეტების იმ მიმდევრობით, რა მიმდევრობითაც მივუთითეთ
    data_columns არგუმენტში.

    არგუმენტები:
        1. file_or_wb_obj  -  Excel-ის ფაილის მისამართი,
                ან შესაბამისი ობიექტი(openpyxl.workbook.workbook.Workbook)

        2. check_column  -  სვეტი, რომლის გამოყენებაც გვინდა ბოლო
                            სტრიქონის იდენტიფიკაციისთვის
                            (სვეტი, რომელშიც ყველაზე მეტია ჩანაწერის
                             არსებობის ალბათობა)

        3. start_row  -  საწყისი სტრიქონის ნომერი,
                            საიდანაც გვინდა მონაცემების მიღება

        4. data_columns -  სვეტები, რომლებიდანაც გვინდა ინფორმაციის მიღება.
                        მაგალითი 1 სიმბოლოიანი სვეტების სახელებისთვის:
                            "ABCD"
                        მაგალითი 2 ან მეტ სიმბოლოიანი სვეტების სახელებისთვის:
                            ["AA", "BB", "AZ"]

        5. sheet_index -  (ნაგულისხმევად = 0, ანუ პირველი ფურცელი)

        6. data_only  -  (ნაგულისხმევად = False) - გამოვიყენოთ True,
                            თუ გვინდა xlsx ფაილში ფორმულების არსებობის
                            შემთხვევაში მათი მნიშვნელობები მივიღოთ
                            ფორმულების ნაცვლად.

        7. number  -  (ნაგულისხმევად = 10), რამდენი ქვედა ცარიელი უჯრა
                    ჩავთვალოთ საკმარისად, რათა უჯრა ბოლოდ მივიჩნიოთ
    '''
    import openpyxl
    # load excel data if workbook object is not directly used
    if not isinstance(file_or_wb_obj, openpyxl.workbook.workbook.Workbook):
        wb = openpyxl.load_workbook(
            file_or_wb_obj, data_only=data_only)
    else:
        wb = file_or_wb_obj

    # Get worksheet
    ws = wb.worksheets[sheet_index]

    # get number of rows we need
    last_row = get_last_row_num(file_or_wb_obj,
                                wb.worksheets[sheet_index].title,
                                check_column,
                                number=number,
                                start_row=start_row)
    # get actual data
    result_list = []

    for row in range(start_row, last_row+1):
        row_records = []

        for column in data_columns:
            row_records.append(str(ws[f'{column}{row}'].value))
        result_list.append(row_records)

    return result_list
