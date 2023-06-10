import docx2txt
from tkinter import Tk, filedialog, Canvas, Button, PhotoImage, Label, StringVar
import spacy
from spacy.matcher import Matcher
import re
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate

nlp = spacy.load('en_core_web_sm')
matcher = Matcher(nlp.vocab)

def extract_text_from_doc():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("DOCX Files", "*.docx")])
    if file_path:
        temp = docx2txt.process(file_path)
        text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
        return ' '.join(text)
    else:
        print("No file selected.")


def extract_name(resume_text):
    nlp_text = nlp(resume_text)
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]
    matcher.add('NAME', [pattern])
    matches = matcher(nlp_text)
    for match_id, start, end in matches:
        span = nlp_text[start:end]
        return span.text


def extract_mobile_number(text):
    phone = re.findall(re.compile(
        r'(?:(?:\+?([1-9]|[0-9][0-9]|[0-9][0-9][0-9])\s*(?:[.-]\s*)?)?(?:\(\s*([2-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9])\s*\)|([0-9][1-9]|[0-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9]))\s*(?:[.-]\s*)?)?([2-9]1[02-9]|[2-9][02-9]1|[2-9][02-9]{2})\s*(?:[.-]\s*)?([0-9]{4})(?:\s*(?:#|x\.?|ext\.?|extension)\s*(\d+))?'),
                       text)
    if phone:
        number = ''.join(phone[0])
        if len(number) > 10:
            return '+' + number
        else:
            return number


def extract_email(email):
    email = re.findall("([^@|\s]+@[^@]+\.[^@|\s]+)", email)
    if email:
        try:
            return email[0].split()[0].strip(';')
        except IndexError:
            return None


def extract_skills(resume_text):
    nlp_text = nlp(resume_text)
    tokens = [token.text for token in nlp_text if not token.is_stop]
    data = pd.read_csv(r'F:\project\skills.csv')
    skills = list(data.columns.values)
    skillset = []
    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)
    for token in nlp_text.noun_chunks:
        token = token.text.lower().strip()
        if token in skills:
            skillset.append(token)
    return [i.capitalize() for i in set([i.lower() for i in skillset])]

def send_email(email, top_jobs):
    quoted_mail = f'{email}'
    sender_email = 'parth.singhal21@pccoepune.org'
    receiver_email = quoted_mail
    subject = 'Email Subject'
    message = "The top jobs for you are\n" + top_jobs[['JOBS']].to_string(index=False)
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    username = 'parth.singhal21@pccoepune.org'
    password = 'Pxrth@8800'
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(username, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print('Email sent successfully!')
        email_sent_label.config(text="Email sent successfully!")  # Update GUI label with success message
    except smtplib.SMTPException as e:
        print('Error sending email:', str(e))
        email_sent_label.config(text="Error sending email")  # Update GUI label with error message


def upload_document():
    resume_text = extract_text_from_doc()
    name = extract_name(resume_text)
    mobile_number = extract_mobile_number(resume_text)
    email = extract_email(resume_text)
    skills_user = extract_skills(resume_text)
    skills_data = pd.read_csv(r'C:\Users\Parth\Desktop\parsight_1\ParSight_Dataset_merged.csv')
    skills_user_str = ' '.join(skills_user)
    vectorizer = TfidfVectorizer()
    skills_job_vectorized = vectorizer.fit_transform(skills_data['SKILLS'])
    skills_user_vectorized = vectorizer.transform([skills_user_str])
    similarities = cosine_similarity(skills_user_vectorized, skills_job_vectorized)[0]
    skills_data['similarity'] = similarities
    sorted_jobs = skills_data.sort_values('similarity', ascending=False)
    num_jobs = 5  # Number of top matching jobs to retrieve
    top_jobs = sorted_jobs.head(num_jobs)

    # Update GUI labels with extracted information
    name_label.config(text="Name: " + name)
    mobile_label.config(text="Mobile number: " + mobile_number)
    email_label.config(text="Email: " + email)
    skills_label.config(text="Skills: " + ", ".join(skills_user))
    table = tabulate(top_jobs[['JOBS', 'similarity']], headers=["Job Title", "Similarity"], showindex=False,tablefmt="presto")
    top_jobs_label.config(text="\t\tTop Matching Jobs\n\n" + table)

    # Send email with top jobs
    send_email(email, top_jobs)


window = Tk()
window.geometry("1290x840")
window.state('zoomed')
#window.attributes('-fullscreen',True)
window.configure(bg="#FFFFFF")

canvas = Canvas(
    window,
    bg="#FFFFFF",
    height=840,
    width=1290,
    bd=0,
    highlightthickness=0,
    relief="ridge"
)
canvas.place(x=0, y=0)

# Rest of your code...
image_image_1 = PhotoImage(file=r"C:\Users\Parth\Desktop\build\assets\frame0\image_1.png")
image_1 = canvas.create_image(400,400, image=image_image_1)

canvas.create_text(
    271.0,
    33.0,
    anchor="nw",
    text="PARSIGHT",
    fill="#FFFFFF",
    font=("Inter SemiBold", 20 * -1)
)

upload_button_image = PhotoImage(file=r"C:\Users\Parth\Desktop\build\assets\frame0\button_1.png")
upload_button = Button(
    image=upload_button_image,
    borderwidth=0,
    highlightthickness=0,
    command=upload_document,
    relief="flat"
)
upload_button.place(x=241.0, y=75.0, width=194.0, height=47.0)

name_label = Label(window, text="Name:", font=("Inter Regular", 12))
name_label.place(x=23.0, y=140.0, anchor="nw")

mobile_label = Label(window, text="Mobile number:", font=("Inter Regular", 12))
mobile_label.place(x=23.0, y=170.0, anchor="nw")

email_label = Label(window, text="Email:", font=("Inter Regular", 12))
email_label.place(x=23.0, y=200.0, anchor="nw")

skills_label = Label(window, text="Skills:", font=("Inter Regular", 12))
skills_label.place(x=23.0, y=230.0, anchor="nw")

top_jobs_label = Label(window, text="Top Matching Jobs:", font=("Inter Regular", 12))
top_jobs_label.place(x=23.0, y=260.0, anchor="nw")

email_sent_label = Label(window, text="", font=("Inter Regular", 12), fg="green")
email_sent_label.place(x=23.0, y=500.0, anchor="nw")


window.mainloop()
