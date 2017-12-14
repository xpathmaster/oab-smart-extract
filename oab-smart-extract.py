import re
import uuid
import shutil
import argparse
import subprocess
import requests

from requests_ntlm import HttpNtlmAuth
requests.packages.urllib3.disable_warnings()

def parse_args():
    
    parser = argparse.ArgumentParser(description='description')
    parser.add_argument('host', metavar='mail_host', type=str, nargs=1,
                        help='mail_host: mail.example.com')
    parser.add_argument('email', metavar='email', type=str, nargs=1,
                        help='email: user@example.com')    
    parser.add_argument('password', metavar='password', type=str, nargs=1,
                        help='pass: password')                            
    
    args = parser.parse_args()  
    return args  

def main():
    args = parse_args()
    host = args.host[0]
    mail_domain = args.email[0].split('@')[1]
    login = args.email[0].split('@')[0]
    password = args.password[0]
  
    # Create a session with ntlm authentication
    session = requests.Session()
    session.auth = HttpNtlmAuth(login, password)

    r = session.get('https://{}/autodiscover/autodiscover.xml'.format(host), verify=False)
    if r.status_code != 200:
        raise Exception('Something went wrong. For example, verify your credentials.')

    r = session.post('https://{}/autodiscover/autodiscover.xml'.format(host), verify=False)

    xml = """
    <Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
      <Request>
        <EMailAddress>{}@{}</EMailAddress>
        <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
      </Request>
    </Autodiscover>""".format(login, mail_domain)

    headers = {'Content-Type': 'text/xml'}
    r = session.post('https://{}/autodiscover/autodiscover.xml'.format(host), data=xml, headers=headers, verify=False)

    found = re.search('<OABUrl>(http.*)</OABUrl>', r.text)

    if found:
        oab_url = found.group(1)
    else:
        raise Exception('OABUrl not found')

    r = session.get(oab_url + "/oab.xml", verify=False)
    found = re.search(r'>(.+lzx)<', r.text)

    if found:
        lzx_link = found.group(1)
    else:
        raise Exception('LZX link not found')

    r = session.get(oab_url + lzx_link, stream=True, verify=False)

    if r.status_code == 200:
        with open("test.lzx", 'wb') as f:
            r.raw.decode_content = True
            shutil.copyfileobj(r.raw, f)

    temp_filename = '/tmp/{}_lzx.txt'.format(uuid.uuid4())
    subprocess.check_output(['oabextract', 'test.lzx', temp_filename])

    file_strings = subprocess.check_output(('strings', temp_filename))
    mails = re.findall(r'[a-zA-Z0-9-_.]+@{}'.format(mail_domain), file_strings)
    unique_mails = sorted(set(mails))
    mails_stringed = '\n'.join(map(str, unique_mails))

    with open("output.txt", "w") as text_file:
        text_file.write(mails_stringed)
    print "DONE. Check file output.txt"
    print "Total emails found:" , len(unique_mails)
if __name__ == '__main__':
    main()
