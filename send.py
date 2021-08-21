from time import sleep
import win32com.client as client
import extract_msg
import tkinter as tk
from tkcalendar import *
from bs4 import BeautifulSoup

username = "careers@hunted.co.il"

root = tk.Tk()
timing_frame = tk.Frame(root, relief= 'sunken', pady=10)
hour_txt = tk.Entry(timing_frame)
cal1 = Calendar(timing_frame, selectmode="none")
e1 = None
subject = ""
email = "careers@hunted.co.il"
sv = tk.StringVar()

def divide_chunks(l, n):
    for i in range(0, len(l), n): 
        yield l[i:i + n]

def send_mails (wait) :
    sleep(int(wait))

    # Convert message.htm file to utf-8
    try :
        from pathlib import Path
        path = Path("message.htm")
        path.write_text(path.read_text(encoding="utf16"), encoding="utf8")

    except Exception :
        pass


    with open('message.htm', 'rb') as f:
        html = f.read()
        html_str = str(html)

        with open('emails.txt', 'r', newline="") as f :
            # reader = csv.reader(f)

            lines = []
            for line in f:
                lines.append(line)

            chunks = list(divide_chunks(lines, 98))

            outlook = client.Dispatch('Outlook.Application')            
            global email
            account = outlook.Session.Accounts[email]
            print(email)

            for chunk in chunks :
                recipients = ""

                for person in chunk :
                    recipients += person.rstrip() + ";"

                # Fix images
                import base64

                soup = BeautifulSoup(html, "html.parser")
                images = soup.findAll('img')

                for image in images :
                    image = image.get('src')
                    
                    if "message_files" in image :
                        print(image)

                        image = image.replace('file:///', '')
                        
                        with open(image, "rb") as image_file :
                            encoded_string = base64.b64encode(image_file.read())
                            
                            if "png" in image :
                                ext = "png"
                            elif "gif" in image :
                                ext = "gif"
                            else :
                                ext = "jpeg"

                            new_pic = 'data:image/' + ext + ';base64, ' + str(encoded_string, 'utf-8')
                            # html = html.replace(bytes(image, 'utf-8'), bytes(new_pic, 'utf-8'))

                message = outlook.CreateItem(0)
                message.BCC = recipients
                message.subject = subject
                message.HTMLbody = html

                try :
                    message._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                except Exception :
                    try:
                        raise TypeError("Email not found")
                    except :
                        pass
                    
                message.Send()

                from datetime import datetime
                now = datetime.now()
                print("sent messages: " + now.strftime("%H:%M:%S"))

                sleep(180)
            print("Finished sending emails.")

            # input("Press Enter to continue...")

def mail_thread():

    date = cal.get_date()
    m = str(min_sb.get()).zfill(2)
    h = str(sec_hour.get()).zfill(2)
    s = str(sec.get()).zfill(2)

    if CheckVar1.get() :
        import time
        import datetime
        from datetime import timezone

        sc_time = date + " " + m + ":" + h + ":" + s
        datetime_object = datetime.datetime.strptime(sc_time, '%d/%m/%Y %H:%M:%S')
        sc_timestamp = datetime_object.replace(tzinfo=timezone.utc).timestamp()

        # Minus 3 hours GMT
        wait = sc_timestamp - time.time() - 10800
        print("wait is: " + str(wait))
    else :
        wait = 0

    root.destroy()
    send_mails(wait)

canvas = tk.Canvas(root, height=0, width=300, bg="#ebebeb")
canvas.pack()

# timing_frame.pack(fill=tk.BOTH, expand= True, padx= 30, pady=20)

def keyup (e) :
    # print(e.__dict__)
    global subject
    subject = e.widget.get("1.0", tk.END)

def keyup_email (e) :
    # print(e.__dict__)
    global email
    email = e.widget.get("1.0", tk.END)

frame = tk.Frame(root, relief= 'sunken')
frame.pack(fill=tk.BOTH, expand= True, padx= 30, pady=20)
e1 = tk.Text(frame, height=1, width=15)
e1.bind("<KeyRelease>", keyup)

e2 = tk.Text(frame, height=1, width=15)
e2.insert(tk.END, email)
e2.bind("<KeyRelease>", keyup_email)

B = tk.Button(frame, text ="Start sending emails", command = mail_thread, pady=10, padx=12, bd=0, bg="#93b5e1")
B.pack()

l1 = tk.Label(frame, text="Subject:")
l1.pack()
e1.pack()

l2 = tk.Label(frame, text="Email:")
l2.pack()
e2.pack()

# Timing options

CheckVar1 = tk.IntVar()

def toggle () :
    if not CheckVar1.get() :
        # timing_frame.pack_forget()
        pass
    else :
        # timing_frame.pack()
        pass

C1 = tk.Checkbutton(frame, command=toggle, text = "Is scheduled?", variable = CheckVar1, onvalue = 1, offvalue = 0, height=5, width = 20)
C1.pack()

cal1.pack()
hour_txt.insert(tk.END, '12:00')
hour_txt.pack()

other = tk.BooleanVar()
tk.Checkbutton(frame, text="Other", variable=other, command=toggle)
ent= tk.Entry(frame,width=50)

hour_string = None
min_string= None
last_value_sec = ""
last_value = ""        
f = ('Times', 20)


if last_value == "59" and min_string.get() == "0":
    hour_string.set(int(hour_string.get())+1 if hour_string.get() !="23" else 0)   
    last_value = min_string.get()

if last_value_sec == "59" and sec_hour.get() == "0":
    min_string.set(int(min_string.get())+1 if min_string.get() !="59" else 0)
if last_value == "59":
    hour_string.set(int(hour_string.get())+1 if hour_string.get() !="23" else 0)            
    last_value_sec = sec_hour.get()

fone = tk.Frame(root)
ftwo = tk.Frame(root)

fone.pack(pady=10)
ftwo.pack(pady=10)

cal = Calendar(fone, selectmode="day", date_pattern="dd/mm/y")
cal.pack()

min_sb = tk.Spinbox(
    ftwo,
    from_=0,
    to=23,
    wrap=True,
    textvariable=hour_string,
    width=2,
    state="readonly",
    font=f,
    justify=tk.CENTER
    )
sec_hour = tk.Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,
    textvariable=min_string,
    font=f,
    width=2,
    justify=tk.CENTER
    )

sec = tk.Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,
    textvariable=sec_hour,
    width=2,
    font=f,
    justify=tk.CENTER
    )

min_sb.pack(side=tk.LEFT, fill=tk.X, expand=True)
sec_hour.pack(side=tk.LEFT, fill=tk.X, expand=True)
sec.pack(side=tk.LEFT, fill=tk.X, expand=True)

msg = tk.Label(
    root, 
    text="Hour  Minute  Seconds",
    font=("Times", 12),
    bg="#93b5e1",
    pady=10,
    padx=10
    )
msg.pack(side=tk.TOP)

tk.mainloop()