import yagmail


def sendemail ():
    try:
        #initializing the server connection
        yag = yagmail.SMTP(user='Mail', password='Password')
        
        yag.send(to=['muhammad.rafi@sabgroupindonesia.com'], subject='Data WN S-26 LP  Report of Gebyar Hadiah Yogya', contents='Yogya Update', attachments=['WN S-26 LP SAB Yogya Gebyar Hebat.xlsx'])
        print("Email sent successfully")
    except:
        print("Error, email was not sent")

