#imports

import os
import re
import subprocess
import pytesseract as tes
from os import listdir
from PIL import Image
from time import sleep
from pathlib import Path

#globals
image_cmd = ['rundll32.exe', r'C:\Program Files\Windows Photo Viewer\PhotoViewer.dll', 'ImageView_Fullscreen']


wdir = Path.home() / "autosort"
sellers_path = wdir / "sellers.txt"

input_path = Path.home()
output_path = Path.home()

incomplete = []
complete = {}
sellers = []

try:
    with open(wdir / "inputpath") as f:
        text = f.read()
        if os.path.exists(text):
            input_path = Path(text)
except IOError:
    pass

try:
    with open(wdir / "outputpath") as f:
        text = f.read()
        if os.path.exists(text):
            output_path = Path(text)
except IOError:
    pass


try:
    with open(sellers_path) as sp:
        for line in sp:
            line = line.strip()
            if len(line) == 0 or line[0] == '#':
                continue
            sellers.append(line)
except IOError:
    f = open(sellers_path, "x")
    f.close

#functions

def get_date(text):
    m = re.search('([0-1]?[0-9])[/-]([0-3]?[0-9])[/-]((?:20)?[0-9][0-9])', text)
    if m == None:
        return None
    strs = m.groups()


    ints = [int(s) for s in strs]

    ints[2] = ints[2] % 100

    return "{}-{}-{}".format(*ints)


def get_seller(text):
    for seller in sellers:
        if text.lower().find(seller.lower()) >= 0:
            return seller

def process_image(im):
    
    text = tes.image_to_string(im, config='--psm 11')

    date = get_date(text)
    seller = get_seller(text)
    

    return {"date": date, "seller": seller }

def name_image(im_meta):
    date_str = im_meta["date"] or "Unknown"

    seller_str = im_meta["seller"] or "Unknown"

    return (seller_str + "_" + date_str)

def process_images():
    print("Attempting to Auto-Sort files.")
    input = [f for f in listdir(input_path) if (input_path / f).is_file()]
    for i, filename in enumerate(input):
        print("Processing {}/{}...".format(i+1, len(input)))
        im = None
        try:
            im = Image.open(input_path / filename)
        except:
            print("couldn't open file {}".format(filename))
            continue
        
        im_meta = process_image(im)
        
        im_meta["filename"] = filename

        im.close()

        if (im_meta["date"] == None) or (im_meta["seller"] == None):
            incomplete.append(im_meta)
            continue
        
        date = im_meta["date"]
        name = name_image(im_meta)

        if name in complete:
            complete[name].append(im_meta)
        else:
            complete[name] = [im_meta]

def save_completed():
    for i, name in enumerate(complete):
        print("saving {}/{}...".format(i+1, len(complete)))
        images = [Image.open(input_path / im_meta["filename"]) for im_meta in complete[name]]

        #saves all the images as a pdf
        images[0].save(
            output_path / (name + ".pdf"), "PDF", resolution=100.0,
            save_all=True, append_images=images[1:]
        )
        for image in images:
            image.close()

def option_input(prompt, opts):
    while True:
        ip = input(prompt).lower().strip()
        if ip in opts:
            return ip
        else:
            print("Invalid Option '{}'".format(ip))

#main

while True:
    print("Enter location of input scans or leave blank to use last ({}):".format(input_path))
    ip = input("\t> ")
    if os.path.exists(ip):
        input_path = Path(ip)
        with open(wdir / "inputpath", "w") as f:
            f.write(ip)
        break
    elif ip == "":
        break
    else:
        print("Invalid Path\n")
        sleep(0.5)

while True:
    print("Enter location to save pdfs to or leave blank to use last ({}):".format(output_path))
    ip = input("\t> ")
    if os.path.exists(ip):
        output_path = Path(ip)
        with open(wdir / "outputpath", "w") as f:
            f.write(ip)
        break
    elif ip == "":
        break
    else:
        print("Invalid Path\n")
        sleep(0.5)

process_images()

if (len(incomplete) > 0):
    input("\nAuto-Sort Complete. Manual Entry Needed for {} files. <Enter> to continue. ".format(len(incomplete)))

while (len(incomplete) > 0):
    im_meta = incomplete[0]
    print("\nFILE MISSING INFORMATION: {}".format(im_meta["filename"]))

    has_seller = (im_meta["seller"] != None)
    has_date = (im_meta["date"] != None)

    viewer = subprocess.Popen(image_cmd + [input_path / im_meta["filename"]])

    if not has_seller:
        im_meta["seller"] = input("\tEnter Seller: ")
    if not has_date:
        date_str = input("\tEnter Date (MM/DD/YY): ")
        im_meta["date"] = get_date(date_str)

    name = name_image(im_meta)

    if (im_meta["seller"] != "") and not (im_meta["seller"] in sellers):
        ip = option_input(
                "{} is not in the list of known sellers. Add it? (y, n) ".format(im_meta["seller"]),
                [ 'y', 'n' ]
        )
        if (ip == "y"):
            new_seller = im_meta["seller"]
            col = False
            for seller in sellers:
                if (seller.lower() in im_meta["seller"].lower()) or (im_meta["seller"].lower() in seller.lower()):
                    print("Seller not added. collision detected with {}".format(seller))
                    col = True
                    sleep(0.5)
                    break;
            if not col:
                sl = open(sellers_path, "a")
                sl.write('\n' + im_meta["seller"])
                sl.close()
                sellers.append(im_meta["seller"])
                print("Added seller.")

    ip = option_input(
        "File will be saved as: {}\n\t<Enter> save and continue\n\t<r> enter different information\n\t<d> discard\n\t<x> exit and discard everything ".format(name),
        [ '', 'r', 'd', 'x' ]
    )

    match ip:
        case "":
            incomplete.pop(0)
            if name in complete:
                complete[name].append(im_meta)
            else:
                complete[name] = [im_meta]
            print("file marked for saving.")
        case "r":
            im_meta["date"] = None
            im_meta["seller"] = None
        case "d":
            incomplete.pop(0)
            print("discarded - file will not be saved to pdf.")
        case "x":
            print("exiting.")
            exit(0)

    viewer.terminate()


save_completed()

ip = option_input(
    "Done saving. Delete raw scans? (y, n) ",
    [ 'y', 'n' ]
)

if ip == 'y':
    for key in complete:
        document = complete[key]
        for page in document:
            try:
                os.remove(input_path / page["filename"])
            except:
                print("couldn't remove {}".format(page["filename"]))

print("\nDone!")
sleep(1.5)
