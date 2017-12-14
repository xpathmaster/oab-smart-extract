# oab-smart-extract
Python script to quickly extract domain emails (addresses) from the Offline Address Book (OAB)

# Installation (Kali Linux):

Install libraries:
```apt install autoconf automake libtool -y```

Install *oabextractor* tool from https://github.com/kyz/libmspack
```
git clone https://github.com/kyz/libmspack
cd libmspack/libmspack/
./rebuild.sh
./configure
make 
make install
```
Install python ntlm-auth support

```
pip install ntlm-auth
```


Download script and use it
```
python oab-smart-extract.py mail_host email password
```

Example:
```
python oab-smart-extract.py mail.example.com user@example.com password
```

All extracted domain emails would be in *output.txt* file
