# coding: utf-8
import subprocess
import sys
import re
import os

class Users():
    """ Creation of the User Class.
    It contains the main information about the user to add in the AD:
    - Name
    - Surname
    - Title
    - Domain name (from the user depends on)
    - Service (from the user depends on)"""

    def __init__(self,**kwargs):

        try:
            self.name = kwargs['name'].capitalize()
        except:
            self.name = ''
        try:
            self.surname = kwargs['surname'].capitalize()
        except:
            self.surname = ''
        try:
            self.title = kwargs['title']
        except:
            self.title = ''
        try:
            self.domain_name = kwargs['domain'].lower()
        except:
            self.domain_name = ''
        try:
            self.service = kwargs['service']
        except:
            self.service = ''

        # Creation of sub-attributes from the main one.
        list_name = list(self.name)
        self.firstletter = list_name[0]
        self.SamAccountName = '{}{}'.format(self.firstletter.lower(),self.surname.lower())
        self.UserPrincipalName = '{}@{}'.format(self.SamAccountName,self.domain_name)
        self.DisplayName = '{} {}'.format(self.name,self.surname)
        list_domain = self.domain_name.split('.')
        self.dc1 = list_domain[0]
        self.dc2 = list_domain[1]


    def SendAD(self,target):
        """ This method sends the powershell command to the AD server in order to add a user with the name, surname and title defined for the instance"""
        cmd = subprocess.Popen(["powershell",'Enter-PSSession {} ; New-ADUser -Name "{}" -Surname "{}" -SamAccountName "{}" -UserPrincipalName "{}" -DisplayName "{}" -ChangePasswordAtLogon 1 -Title "{}" -Path "OU=Users,OU={},OU=Services,DC={},DC={}" -AccountPassword(ConvertTo-SecureString p@ssw0rd% -AsPlainText -Force) -Enabled $true ; exit'.format(target, self.name,self.surname,self.SamAccountName,self.UserPrincipalName,self.DisplayName,self.title,self.service,self.dc1,self.dc2)])
        cmd.terminate()
