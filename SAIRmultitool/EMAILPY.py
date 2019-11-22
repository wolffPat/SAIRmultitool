import os
import sys
import win32com.client as win32


def Emailer(text, subject, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.CC = "vuit.incident.response@vanderbilt.edu"
    mail.HtmlBody = text
    mail.Display(False)


Emailer("Collaboration On-Call: <br><br>This phish response request should execute all requested functions on email systems both on-premises (standard Exchange) and the cloud (any accounts migrated to Office365). <br><br>Purge requested if multiple recipients. <br><br>In reference to:  <paste IM# from Communication Ticket that will be opened up in next step> <br><br>Please respond with the time/date the sender was blocked. If the sender is an internal sender that is confirmed to be spamming please let us know ASAP.  <br><br>Please add the following URL to the APT block list: (If the link is blocked by ATP, you do not need to add the URL to the block list)<If the phishing URL was wrapped by ATP, include the decoded link (http://www.o365atp.com/) ex.na01.safelinks.protection.outlook.com/?url > <br><br>Once the sender has been blocked, please provide what information you can with this phishing email (please see attached): <br><br>•         How many received by Vanderbilt? <br>•         How many blocked at email hold-points?  <br>•         How many delivered to internal recipients? <br>•         How many people have replied back? <br>•         List of recipients? <br>•         How many people replied back to the Phish?  <br>•         When the first arrived? <br>" , '' , 'vuit.collaboration@vanderbilt.edu')
