import sys
from helper import *
def main():
    try:
        excel_path = sys.argv[1]
        word_path = sys.argv[2]
    except IndexError:
        raise SystemExit("add 2 paths as arguments: path to an excel spreadsheet and path to a word doc template")
    to_list, cc_list, attachments_list, keywords_list = read_excel_data(excel_path)
    subject, text = read_word_data(word_path)
    for idx in range(len(to_list)):
        to, cc, attachments = to_list[idx], None, None
        if cc_list is not None:
            cc = cc_list[idx]
        if attachments_list is not None:
            attachments = attachments_list[idx]
        keywords = {key: keywords_list[key][idx] for key in keywords_list}
        body = substitute_data(text, keywords)
        email, password, smtp_server, smtp_port = email_credentials()
        send_email(smtp_server, smtp_port, email, password, to, cc, subject, body, attachments)
if __name__ == "__main__":
    main()
