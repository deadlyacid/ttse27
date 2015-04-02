import xlwt

__author__ = 'elad'
# ************************************************************************
# This is the second version of SpeechRecognition of Idomoo-Audio files
# To run the script follow the instruction by run:
# ttse.py -h
# or
# ttse.py --help
#
#
# Last updated: 22/03/2015
# *************************************************************************


import datetime
import sys
import getopt
import xlrd
import os
import wave
import contextlib
import speech_recognition as sr
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf


def main():
    user_args = get_args()
    data_from_xls = read_file(user_args)
    results = speech_recognition(data_from_xls, user_args)
    write_results(user_args, results)


def speech_recognition(data_from_xls, user_args):
    base_path = user_args[1]
    excel_file = user_args[0]
    language = user_args[2]

    r = sr.Recognizer()
    results = []

    for voice_file, scene_sentence, sheet_name in data_from_xls:
        print(sheet_name)
        if isinstance(voice_file, float):
            voice_file = str(int(voice_file))
        else:
            voice_file = str(voice_file)

        if ".wav" not in voice_file:
            voice_file += ".wav"

        curr_file = os.path.join(base_path, voice_file)
        print("Trying to load file: %s" % curr_file)
        with sr.WavFile(curr_file) as source:
            r.language = language  # use "test.wav" as the audio source
            audio = r.record(source)  # extract audio data from the file
            print("In Speech-Recognition of: %s" % curr_file)
            try:
                output = r.recognize(audio)
                duration_in_sec = get_file_duration(curr_file)
                results.append((voice_file, output, scene_sentence, duration_in_sec, sheet_name))
                # print(results)
            except LookupError as x:
                duration_in_sec = 'NULL'
                results.append((voice_file, "ERR IN RECOGNITION", scene_sentence, duration_in_sec, sheet_name))
                print 'Lookup error in recognize: %s -  %s' % (str(x), str(locals()))

    print "results:::"
    print results
    return results


def get_file_duration(curr_file):
    fname = curr_file
    with contextlib.closing(wave.open(fname, 'r')) as f:
        frames = f.getnframes()
        rate = f.getframerate()
        duration = frames / float(rate)
        duration = "{:10.2f}".format(duration)

    return duration


def write_results(user_args, results):
    # print(results)
    language = user_args[2]
    rb = xlrd.open_workbook(user_args[0])
    # a writable copy (I can't read values out of this, only write to it)
    wb = copy(rb)

    curr_sheet_name = ""
    check_en_diff = True
    if str(language).lower() != 'en':
        check_en_diff = False
    for idx, tp in enumerate(results):
        filename, result_from_sr, scene_sentence, sound_duration, sheet_name = tp

        if sheet_name != curr_sheet_name:
            my_idx = 0
            curr_sheet_name = sheet_name

            w_sheet = wb.add_sheet(curr_sheet_name + '_results')
            # w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy
            headlines_style = easyxf('font: bold on;')
            w_sheet.write(0, 0,  "File Name", headlines_style)
            w_sheet.write(0, 1, "Scene Sentence", headlines_style)
            w_sheet.write(0, 2, "SpeechRecognition Result", headlines_style)
            w_sheet.write(0, 3, "Sound Duration", headlines_style)

            w_sheet.col(0).width = 256*20
            w_sheet.col(1).width = 256*60
            w_sheet.col(2).width = 256*60
            w_sheet.col(3).width = 256*15
        if check_en_diff:
            if str(result_from_sr).lower() != str(scene_sentence).lower():
                st = easyxf('pattern: pattern solid;')
                st.pattern.pattern_fore_colour = 2
            else:
                st = easyxf('pattern: pattern solid;')
                st.pattern.pattern_fore_colour = xlwt.Style.colour_map['white']
        else:
            st = easyxf('pattern: pattern solid;')
            st.pattern.pattern_fore_colour = xlwt.Style.colour_map['white']
        w_sheet.write(my_idx + 1, 0, filename)
        w_sheet.write(my_idx + 1, 1, scene_sentence)
        w_sheet.write(my_idx + 1, 2, result_from_sr, st)
        w_sheet.write(my_idx + 1, 3, sound_duration)
        my_idx += 1
    # for row_index in range(START_ROW, r_sheet.nrows):
    #     age_nov = r_sheet.cell(row_index, col_age_november).value
    #     if age_nov == 3:
    #         #If 3, then Combo I 3-4 year old  for both summer1 and fall1
    #         w_sheet.write(row_index, col_summer1, 'Combo I 3-4 year old')
    #         w_sheet.write(row_index, col_fall1, 'Combo I 3-4 year old')
    #

    curr_datetime = str(datetime.datetime.now().strftime("%Y-%m-%d-%H-%M").replace("-", ""))
    output_file = user_args[0]
    if ".xls" or ".xlsx" in user_args[0]:
        output_file = str(user_args[0].split(".xls")[0])

    output_file += "_" + curr_datetime + ".xls"
    wb.save(output_file)


def read_file(user_args):
    if len(user_args[0]) < 3:
        raise Exception("Input file is not valid!")
        sys.exit(2)
    data_from_xls = []
    book = xlrd.open_workbook(user_args[0])

    for sheet in user_args[3]:
        curr_sheet = book.sheet_by_index(int(sheet)-1)
        # returns array of tupples with (file_name, text, what it should be
        data_from_xls += map(lambda x, y: (x, y, curr_sheet.name), curr_sheet.col_values(0)[1:], curr_sheet.col_values(1)[1:])
        # for x in range(1, number_of_rows):
        #     if isinstance(curr_sheet.col_values(0)[x], float):
        #         curr_coloumn_val = str(int(curr_sheet.col_values(0)[x]))
        #     else:
        #         curr_coloumn_val = str(curr_sheet.col_values(0)[x])
        #
        #     voice_files_from_xls.append(curr_coloumn_val)
    #print(voice_files_from_xls)

    return data_from_xls


def get_args():
    inputfile = ''
    basepath = ''
    language = 'en'
    selected_sheet = 0
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:l:b:s:",
                                   ["help", "input_file=", "language=", "base_path=", "selected_sheet="])
    except getopt.GetoptError:
        print('ttse.py -i <inputfile> -b <base_path> -l <language> -s <selected_sheet; ex. 1,3,8>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('ttse.py -i <inputfile> -b <base_path> -l <language> -s <selected_sheet; ex. 1,3,8>')
            sys.exit()
        elif opt in ("-i", "--input_file"):
            inputfile = arg
        elif opt in ("-b", "--base_path"):
            basepath = arg
        elif opt in ("-l", "--language"):
            language = arg
        elif opt in ("-s", "--selected_sheet"):
            selected_sheet = arg

    if hasattr(selected_sheet, "__len__") is False:
        selected_sheet = [str(int(selected_sheet))]
    else:
        selected_sheet = selected_sheet.split(",")
    # tuples are like final array in java
    user_args = (inputfile, basepath, language, selected_sheet)

    for user_inputs in user_args:
        # print(len(user_inputs))
        if len(user_inputs) == 0:
            print("Some required args have not been typed.\nPlease follow the instructions:\nttse.py -help")
            exit(2)

    # print(user_inputs)
    return user_args


if __name__ == "__main__":
    main()