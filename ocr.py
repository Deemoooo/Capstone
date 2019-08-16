import locale
import time
import pytesseract
import pandas as pd
import cv2
import tesserocr as tr
import re
import os
from PIL import Image
from pdf2image import convert_from_path
from shutil import rmtree


class run_ocr:

    def __init__(self, pdf, doctype, printer):
        self._pdf = pdf
        self._doctype = doctype
        self._printer = printer

    def get_num_page(self):
        '''
        doctype: BaS, BiS, LS
        '''
        pages = convert_from_path(self._pdf)  # CHANGE THIS

        index = 1
        # file path to store the images converted from input PDF
        img_path = 'sample' + self._doctype

        if img_path in os.listdir():
            rmtree(img_path)
        os.mkdir(img_path)

        for page in pages:
            pathlok = img_path + "/" + str(index) + ".jpg"
            page.save(pathlok, 'JPEG')
            img = cv2.imread(pathlok, 1)
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

            '''
            Custom for Citi BS
            '''
            if self._doctype == 'BaS':
                if index == 1:
                    # 850-2000 is the range that transaction details appears on the first page
                    # values are calculated wrt bank statement from Citibank
                    crop = img[850:2000, :]
                    cv2.imwrite(pathlok, crop)
                else:
                    # 370-2000 is the range that transaction details appears on the rest of the pages
                    # values are calculated wrt bank statement from Citibank
                    crop = img[370:2000, :]
                    cv2.imwrite(pathlok, crop)
                index += 1
            elif self._doctype == 'LR':
                img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                hei, wid = img.shape
                # belwo cropping is the range that contains loan repayment details
                # values are calculated wrt loan repayment document from DBS
                crop = img[int(0.34091 * hei):int(0.66061 * hei), int(0.19608 * wid):]
                cv2.imwrite(pathlok, crop)
                index += 1
            elif self._doctype == 'BiS':
                img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                hei, wid = img.shape
                # belwo cropping is the range that contains bill statement details
                # values are calculated wrt bill statement from DBS
                crop = img[int(0.26818 * hei):int(0.87045 * hei), int(0.21471 * wid):int(0.86471 * wid)]
                cv2.imwrite(pathlok, crop)
                index += 1

        all_file = os.listdir(img_path)
        num_page = len(all_file)

        return num_page

    def combiner(self, column_ls):
        y_list = [i[1] for i in column_ls]
        try:
            dist = [j - i for i, j in zip(y_list[:-1], y_list[1:])]
            count = sum(i < 5 for i in dist)
            dist_bool = [i < 10 for i in dist]
        except:
            count = 0

        if count < 5:
            return column_ls

        idx = 0
        true_column = []
        for i in range(len(column_ls)):
            try:
                if dist_bool[i]:
                    true_column.append([])
                true_column[idx].append(column_ls[i])
                if not dist_bool[i]:
                    idx += 1
            except:
                true_column[idx].append(column_ls[i])

        for i in range(len(true_column)):
            true_column[i] = sorted(true_column[i], key=lambda x: x[0])

        result = [(i[0][0],
                   i[0][1],
                   i[len(i) - 1][0] + i[len(i) - 1][2] - int(i[0][0]),
                   i[len(i) - 1][1] + i[len(i) - 1][3] - int(i[0][1]),
                   int(sum([j[4] for j in i]) / len(i)),
                   ''.join([str(j[5]) for j in i])) for i in true_column]

        return result

    def positioner(self, data, coor, conf, column_ls, column_name, y_sep, line_dist, datatype='int'):
        '''
        Update all dataframe by positioning values into their correct
        location. The search is location based and only works for single
        line entries.
        '''
        for i in column_ls:
            if datatype == 'int':
                char = re.sub("[^0-9]", '', i[5])
            elif datatype == 'str':
                char = i[5]
            cor = ' '.join([str(j) for j in i[:4]])
            if char != '':
                if datatype == 'int':
                    char = float(char[:-2] + '.' + char[-2:])
                for j in range(1, len(y_sep)):
                    # 0.5*line_dist is a buffer value to measure the coordinates
                    if i[1] < y_sep[j] - 0.5 * line_dist:
                        ' '.join([str(j) for j in i[:4]])
                        try:
                            data.at[j - 1, column_name] = char
                        except:
                            break
                        coor.at[j - 1, column_name] = cor
                        conf.at[j - 1, column_name] = i[4]
                        break
                    elif j == len(y_sep) - 1:
                        try:
                            data.at[j, column_name] = char
                        except:
                            break
                        coor.at[j, column_name] = cor
                        conf.at[j, column_name] = i[4]

    def positioner_selector(self, data, coor, conf, column_ls, column_name, y_sep, line_dist, identifier):
        '''
        Update all dataframe by positioning values into their correct
        location. The search is based on both location and identifier.
        Works for multiline entries.
        '''
        for i in column_ls:
            cor = ' '.join([str(j) for j in i[:4]])
            if identifier[column_name] in i[5]:
                char = re.sub('^' + identifier[column_name] + '|\.$', '', i[5])
                for j in range(1, len(y_sep)):
                    if i[1] < y_sep[j] - 0.5 * line_dist:
                        ' '.join([str(j) for j in i[:4]])
                        data.at[j - 1, column_name] = char
                        coor.at[j - 1, column_name] = cor
                        conf.at[j - 1, column_name] = i[4]
                        break
                    elif j == len(y_sep) - 1:
                        data.at[j, column_name] = char
                        coor.at[j, column_name] = cor
                        conf.at[j, column_name] = i[4]

    def single_page_bas(self, image, page_number):
        cv_img = cv2.imread(image)
        hei, wid, _ = cv_img.shape
        pil_img = Image.open(image)
        api = tr.PyTessBaseAPI(psm=tr.PSM.SPARSE_TEXT)
        api.SetImage(pil_img)
        boxes = api.GetComponentImages(tr.RIL.TEXTLINE, True, raw_image=False)
        # print(boxes)
        # text = api.GetUTF8Text()
        # print(text)
        # e is a buffer value to make the bounding box more accurate
        e = 4.5

        # measured x-coordinate for each column of information
        x_sep = [172, 720, 950, 1180]
        column = [[] for i in range(len(x_sep) + 1)]

        '''
        Separate column entries based on x-coordinate.
        '''
        for (im, box, _, _) in boxes:
            x, y, w, h = box['x'], box['y'], box['w'], box['h']
            custom_x = max(x - e, 0)
            custom_y = max(y - e, 0)
            custom_w = min(w + 2 * e, wid - custom_x)
            custom_h = min(h + 2 * e, hei - custom_y)
            api.SetRectangle(custom_x, custom_y, custom_w, custom_h)
            text = api.GetUTF8Text()
            conf = api.MeanTextConf()
            result = (x, y, w, h, conf, text.rstrip())
            if x < x_sep[0]:
                column[0].append(result)
            elif x < x_sep[1]:
                column[1].append(result)
            elif x < x_sep[2]:
                column[2].append(result)
            elif x < x_sep[3]:
                column[3].append(result)
            else:
                column[4].append(result)

        data = pd.DataFrame({'Date': pd.Series([], dtype=str),
                             'Debit': pd.Series([], dtype=float),
                             'Credit': pd.Series([], dtype=float),
                             'Balance': pd.Series([], dtype=float),
                             'Info': pd.Series([], dtype=str),
                             'Payer': pd.Series([], dtype=str),
                             'Payee': pd.Series([], dtype=str)
                             })
        coor = pd.DataFrame({'Date': [],
                             'Debit': [],
                             'Credit': [],
                             'Balance': [],
                             'Info': [],
                             'Payer': [],
                             'Payee': [],
                             'Page Number': []
                             }, dtype=str)
        conf = pd.DataFrame({'Date': [],
                             'Debit': [],
                             'Credit': [],
                             'Balance': [],
                             'Info': [],
                             'Payer': [],
                             'Payee': []
                             }, dtype=int)

        data['Date'] = [i[5] for i in column[0]]
        coor['Date'] = [' '.join([str(j) for j in i[:4]]) for i in column[0]]
        conf['Date'] = [float(i[4]) for i in column[0]]
        coor['Page Number'] = page_number

        y_sep = [i[1] for i in column[0]]
        # print(y_sep)
        try:
            line_dist = min([j - i for i, j in zip(y_sep[:-1], y_sep[1:])])
        except:
            line_dist = 0

        self.positioner(data, coor, conf, column[2], 'Debit', y_sep, line_dist)
        self.positioner(data, coor, conf, column[3], 'Credit', y_sep, line_dist)
        self.positioner(data, coor, conf, column[4], 'Balance', y_sep, line_dist)

        identifier = {'Payee': 'B/O PAYONEER ', 'Payer': 'BENEFICIARY : '}

        self.positioner_selector(data, coor, conf, column[1], 'Payer', y_sep, line_dist, identifier)
        self.positioner_selector(data, coor, conf, column[1], 'Payee', y_sep, line_dist, identifier)

        '''
        Citi bank statement unique workflow. Locate the transaction info corresponding
        to each date value.
        '''
        for i in range(len(y_sep)):
            diff = [abs(y_sep[i] - float(k[1])) for k in column[1]]
            idx = diff.index(min(diff))
            res = column[1][idx]
            char = re.sub('\.$', '', res[5])
            cor = ' '.join([str(j) for j in res[:4]])
            data.at[i, 'Info'] = char
            coor.at[i, 'Info'] = cor
            conf.at[i, 'Info'] = res[4]

        api.End()

        del_idx = data.index[data['Date'] == ''].tolist()
        data.drop(del_idx, inplace=True)
        coor.drop(del_idx, inplace=True)
        conf.drop(del_idx, inplace=True)

        return data, coor, conf

    def single_page_lr(self, image, page_number):
        cv_img = cv2.imread(image)
        cv_img = cv2.medianBlur(cv_img, 5)
        hei, wid, _ = cv_img.shape
        pil_img = Image.fromarray(cv_img)
        api = tr.PyTessBaseAPI(psm=tr.PSM.SPARSE_TEXT_OSD)
        api.SetImage(pil_img)
        boxes = api.GetComponentImages(tr.RIL.TEXTLINE, True)

        # print(boxes)
        # text = api.GetUTF8Text()
        # print(text)

        # e is a buffer value to make the bounding box more accurate
        e = 20

        # measured x-coordinate for each column of information
        x_sep = [0.11765 * wid, 0.27059 * wid, 0.39216 * wid, 0.54902 * wid]
        column = [[] for i in range(len(x_sep) + 1)]
        '''
        Separate column entries based on x-coordinate.
        '''
        for (im, box, _, _) in boxes:
            x, y, w, h = box['x'], box['y'], box['w'], box['h']
            custom_x = max(x - e, 0)
            custom_y = max(y - e, 0)
            custom_w = min(w + 2 * e, wid - custom_x)
            custom_h = min(h + 2 * e, hei - custom_y)
            api.SetRectangle(custom_x, custom_y, custom_w, custom_h)
            text = api.GetUTF8Text()
            conf = api.MeanTextConf()
            result = (x, y, w, h, conf, text.rstrip())
            if x < x_sep[0]:
                column[0].append(result)
            elif x < x_sep[1]:
                column[1].append(result)
            elif x < x_sep[2]:
                column[2].append(result)
            elif x < x_sep[3]:
                column[3].append(result)
            else:
                column[4].append(result)

        data = pd.DataFrame({'Due Date': pd.Series([], dtype=str),
                             'Repayment Amount': pd.Series([], dtype=float),
                             'Principal Payment': pd.Series([], dtype=float),
                             'Interest Payment': pd.Series([], dtype=float),
                             'Outstanding Balance': pd.Series([], dtype=float)
                             })
        coor = pd.DataFrame({'Due Date': [],
                             'Repayment Amount': [],
                             'Principal Payment': [],
                             'Interest Payment': [],
                             'Outstanding Balance': []
                             }, dtype=str)
        conf = pd.DataFrame({'Due Date': [],
                             'Repayment Amount': [],
                             'Principal Payment': [],
                             'Interest Payment': [],
                             'Outstanding Balance': []
                             }, dtype=int)

        data['Due Date'] = [re.sub("[^0-9//]", '', i[5]) for i in column[0]]
        coor['Due Date'] = [' '.join([str(j) for j in i[:4]]) for i in column[0]]
        conf['Due Date'] = [float(i[4]) for i in column[0]]
        coor['Page Number'] = page_number

        y_sep = [i[1] for i in column[0]]
        # print(y_sep)
        try:
            line_dist = min([j - i for i, j in zip(y_sep[:-1], y_sep[1:])])
        except:
            line_dist = 0

        self.positioner(data, coor, conf, column[1], 'Repayment Amount', y_sep, line_dist)
        self.positioner(data, coor, conf, column[2], 'Principal Payment', y_sep, line_dist)
        self.positioner(data, coor, conf, column[3], 'Interest Payment', y_sep, line_dist)
        self.positioner(data, coor, conf, column[4], 'Outstanding Balance', y_sep, line_dist)

        api.End()

        del_idx = data.index[data['Due Date'] == ''].tolist()
        data.drop(del_idx, inplace=True)
        coor.drop(del_idx, inplace=True)
        conf.drop(del_idx, inplace=True)

        return data, coor, conf

    def single_page_bis(self, image, page_number):
        cv_img = cv2.imread(image)
        # cv_img = cv2.medianBlur(cv_img, 5)
        hei, wid, _ = cv_img.shape
        pil_img = Image.fromarray(cv_img)
        api = tr.PyTessBaseAPI(psm=tr.PSM.SPARSE_TEXT_OSD)
        api.SetImage(pil_img)
        boxes = api.GetComponentImages(tr.RIL.TEXTLINE, True, raw_padding=100)
        # print(boxes)
        # text = api.GetUTF8Text()
        # print(text)

        # e is a buffer value to make the bounding box more accurate
        e = 16

        # measured x-coordinate for each column of information
        x_sep = [0.23529 * wid, 0.33031 * wid, 0.40723 * wid, 0.70135 * wid, 0.84615 * wid]
        column = [[] for i in range(len(x_sep) + 1)]

        '''
        Separate column entries based on x-coordinate.
        '''
        for (im, box, _, _) in boxes:
            x, y, w, h = box['x'], box['y'], box['w'], box['h']
            custom_x = max(x - e, 0)
            custom_y = max(y - e, 0)
            custom_w = min(w + 2 * e, wid - custom_x)
            custom_h = min(h + 2 * e, hei - custom_y)
            api.SetRectangle(custom_x, custom_y, custom_w, custom_h)
            text = api.GetUTF8Text()
            conf = api.MeanTextConf()
            result = (x, y, w, h, conf, text.rstrip())
            if x < x_sep[0]:
                column[0].append(result)
            elif x < x_sep[1]:
                column[1].append(result)
            elif x < x_sep[2]:
                column[2].append(result)
            elif x < x_sep[3]:
                column[3].append(result)
            elif x < x_sep[4]:
                column[4].append(result)
            else:
                column[5].append(result)

        data = pd.DataFrame({'Deal No.': pd.Series([], dtype=str),
                             'Item No.': pd.Series([], dtype=str),
                             'Currency': pd.Series([], dtype=str),
                             'Outstanding Balance': pd.Series([], dtype=float),
                             'Trans. Date': pd.Series([], dtype=str),
                             'Expiry Date': pd.Series([], dtype=str)
                             })
        coor = pd.DataFrame({'Deal No.': [],
                             'Item No.': [],
                             'Currency': [],
                             'Outstanding Balance': [],
                             'Trans. Date': [],
                             'Expiry Date': [],
                             'Page Number': []
                             }, dtype=str)
        conf = pd.DataFrame({'Deal No.': [],
                             'Item No.': [],
                             'Currency': [],
                             'Outstanding Balance': [],
                             'Trans. Date': [],
                             'Expiry Date': []
                             }, dtype=int)

        data['Deal No.'] = [re.sub("[^0-9//]", '', i[5]) for i in column[0]]
        coor['Deal No.'] = [' '.join([str(j) for j in i[:4]]) for i in column[0]]
        conf['Deal No.'] = [float(i[4]) for i in column[0]]
        coor['Page Number'] = page_number

        y_sep = [i[1] for i in column[0]]
        # print(y_sep)
        try:
            line_dist = min([j - i for i, j in zip(y_sep[:-1], y_sep[1:])])
        except:
            line_dist = 0

        self.positioner(data, coor, conf, column[1], 'Item No.', y_sep, line_dist, datatype='str')
        self.positioner(data, coor, conf, column[2], 'Currency', y_sep, line_dist, datatype='str')
        self.positioner(data, coor, conf, self.combiner(column[3]), 'Outstanding Balance', y_sep, line_dist)
        self.positioner(data, coor, conf, column[4], 'Trans. Date', y_sep, line_dist, datatype='str')
        self.positioner(data, coor, conf, column[5], 'Expiry Date', y_sep, line_dist, datatype='str')

        api.End()

        del_idx = data.index[data['Deal No.'] == ''].tolist()
        data.drop(del_idx, inplace=True)
        coor.drop(del_idx, inplace=True)
        conf.drop(del_idx, inplace=True)

        return data, coor, conf

    def multi_page_reader(self):
        if self._doctype == 'BaS':
            print('[OCR] Starting OCR...')
            print('[OCR] Reading Bank Statement...')
            docname = 'Bank Statement'
        elif self._doctype == 'BiS':
            print('[OCR] Starting OCR...')
            print('[OCR] Reading Bill Statement...')
            docname = 'Bill Statement'
        elif self._doctype == 'LR':
            print('[OCR] Starting OCR...')
            print('[OCR] Reading Loan Repayment Schedule...')
            docname = 'Loan Repayment Schedule'
        else:
            print('[OCR] Error: Invalid Doctype!')
            return ""
        num_page = self.get_num_page()
        start = time.time()
        self._printer.sprint('Reading {} ({} pages)...'.format(docname, num_page), 0, 0)
        print('[OCR] Progress: 0/{} (0%) Time Elapsed: 0.00s'.format(num_page))
        data = None
        coor = None
        conf = None
        for i in range(num_page):
            page_number = i + 1

            progress = int(page_number / num_page * 100)
            page_path = 'sample{}/{}.jpg'.format(self._doctype, page_number)
            if self._doctype == 'BaS':
                dat, coo, con = self.single_page_bas(page_path, page_number)
            elif self._doctype == 'BiS':
                dat, coo, con = self.single_page_bis(page_path, page_number)
            else:
                dat, coo, con = self.single_page_lr(page_path, page_number)
            data = pd.concat([data, dat], ignore_index=True)
            coor = pd.concat([coor, coo], ignore_index=True)
            conf = pd.concat([conf, con], ignore_index=True)
            end = time.time()
            elapsed = end - start
            print('[OCR] Progress: {}/{} ({}%) Time Elapsed: {:.2f}s'.format(page_number, num_page, progress, elapsed))
            self._printer.sprint('({}) Progress: {}/{} Completed'.format(docname, page_number, num_page), 2, 0)
        print('[OCR] The time has come!')
        self._printer.sprint('Reading {} ({} pages)...Done'.format(docname, num_page), 0, 0)

        return data, coor, conf