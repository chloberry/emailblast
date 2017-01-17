import smtplib
import getpass

from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

print ("Thanks for using emailblast! Built by Chloe Chan.\n")
my_address = input("Please type in your email: ")
password = getpass.getpass('Password for %s: ' % my_address)


# Returns list from 'filename' containing names and corresponding emails
def get_contacts():

    names = []
    emails = []
    with open('mycontacts.txt', mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails


# Returns the content inside template object of 'filename'
def read_template():

    with open('template.txt', 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)



def main():
    names, emails = get_contacts()  # Read contacts
    message_template = read_template()  # Read template

    # Set up the SMTP server
    server = smtplib.SMTP('smtp.office365.com', '587')
    server.starttls()
    server.login(my_address, password)

    # Displays stylized email signature in HTML
    html = """\
    <html>
        <body>
            <p class="p1"><span class="s1">....</span></p>
<p class="p2">&nbsp;</p>
<p class="p2">&nbsp;</p>
<p class="p3"><span class="s1"><b>C H L O E &nbsp; &nbsp;&nbsp;C H A N&nbsp;</b></span></p>
<p class="p4"><span class="s1"><i>Queen&rsquo;s University BComm &lsquo;17</i></span></p>
<p class="p5">&nbsp;</p>
<p class="p6"><span class="s1">Senior Advisor&nbsp;</span><span class="s2">&nbsp;|</span><span class="s1">&nbsp;&nbsp;TEDxQueensU</span></p>
<p class="p6">Teaching Assistant | Smith School of Business<br />
&nbsp;</p>
<p class="p7"><span class="s3"><a href="mailto:chloe.chan@queensu.ca">chloe.chan@queensu.ca</a></span></p>
<p class="p6"><span class="s1">647.501.2881</span></p>
<p class="p6"><a href="http://linkedin.com/in/chloechanto" style="-webkit-text-stroke-color: rgb(0, 121, 205);">LinkedIn</a></p>
<style type="text/css">
p.p1 {margin: 0.0px 0.0px 0.0px 0.0px; font: 12.0px Helvetica; -webkit-text-stroke: #000000}
p.p2 {margin: 0.0px 0.0px 0.0px 0.0px; font: 10.0px Arial; -webkit-text-stroke: #000000; min-height: 11.0px}
p.p3 {margin: 0.0px 0.0px 0.0px 0.0px; font: 12.0px Arial; -webkit-text-stroke: #000000}
p.p4 {margin: 0.0px 0.0px 0.0px 0.0px; font: 10.0px Arial; -webkit-text-stroke: #000000}
p.p5 {margin: 0.0px 0.0px 0.0px 0.0px; font: 9.0px Arial; -webkit-text-stroke: #000000; min-height: 10.0px}
p.p6 {margin: 0.0px 0.0px 0.0px 0.0px; font: 10.0px Arial; color: #b5b5b5; -webkit-text-stroke: #b5b5b5}
p.p7 {margin: 0.0px 0.0px 0.0px 0.0px; font: 10.0px Arial; color: #4787ff; -webkit-text-stroke: #4787ff}
p.p8 {margin: 0.0px 0.0px 0.0px 0.0px; font: 10.0px Arial; color: #0079cd; -webkit-text-stroke: #0079cd}
span.s1 {font-kerning: none}
span.s2 {font-kerning: none; color: #e0e0e0; -webkit-text-stroke: 0px #e0e0e0}
span.s3 {text-decoration: underline ; font-kerning: none}</style>
 </p>
        </body>
    </html>
    """



    # For each contact, send the email:
    for name, email in zip(names, emails):
        msg = MIMEMultipart()  # create a message

        # Add in the actual person name to the message template
        message = message_template.substitute(PERSONNAME=name.title())

        # Prints out the message body for our sake
        print(message)

        # Set up the parameters of the message
        msg['From'] = my_address
        msg['To'] = email
        msg['Subject'] = "Next Generation Athletic Fields"

        # Add in the message body
        msg.attach(MIMEText(message, 'plain'))
        msg.attach(MIMEText(html, 'html'))

        # Send the message via the server set up earlier.
        server.sendmail(my_address, email, msg.as_string())
        del msg

    # Terminate the SMTP session and close the connection
    server.quit()


if __name__ == '__main__':
    main()