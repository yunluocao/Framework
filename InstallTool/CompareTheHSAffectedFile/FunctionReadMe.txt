1.Windows Installation:
(1)Usually windows python has the pip tool after install python, if no, please download the setup.py, and using command"python setup.py install"
Then add the Path environment parameter in the C:\Python27\Scripts
(2)Using the command pip install to install openpyxl and bs4, ex: pip install bs4



2.Linux Installation:
(1)Install pip from the below steps:
<1>wget https://bootstrap.pypa.io/get-pip.py(If that wget not found, please use the command:yum -y install wget）
<2>python get-pip.py
<3>pip -V#check it has installed successfully.

Note:
As the installed dir may have the authorization issue, please use sudo in the frone of the command.
After installed successfully, please modify the installed dir permission,so no need to use sudo later,
chmod 777 /usr/lib/python2.7/site-packages
[hci@CNSHVLCENT01 qdxinstall]$ pip -V
pip 9.0.1 from /usr/lib/python2.7/site-packages (python 2.7)

(2)Using the command pip install to install openpyxl and bs4, ex: pip install bs4



3.Unix Installation
(1)Install pip from the below steps:
As some pip version isn't applicable for the unix, it's better to installed the setup.py to install.
<1>Find one location, and use the below command to download the setup.py
wget "https://pypi.python.org/packages/source/p/pip/pip-1.5.4.tar.gz#md5=834b2904f92d46aaa333267fb1c922bb" --no-check-certificate
<2>Use the gunzip command to unzip the gz package, then use the command tar -xvf to unzip the tar package
tar -axf pip-1.5.4.tar.gz
<3>cd the $downloadDir/pip-1.5.4/
<4>python setup.py install
Note:the permission issue it could refer the linux part.
After installation, pip -V may report error: ksh:pip: command not found...

<5>As the command may haven't added to the /user/bin,it could use the softlink to handle it(It must use softlink, not hardlink, as softlink could handle base
on the different file system)
<6>It could use the find command to find the pip install dir:find / -name pip
ex:
$ sudo find / -name pip
/home/hci/pip-1.5.4/pip-1.5.4/build/lib/pip
/home/hci/pip-1.5.4/pip-1.5.4/pip
/opt/freeware/bin/pip
/opt/freeware/lib/python2.7/site-packages/pip
/opt/freeware/lib/python2.7/site-packages/pip-1.5.4-py2.7.egg/pip
find: bad status-- /proc/6553732/fd/5
find: bad status-- /proc/6553732/fd/6
find: bad status-- /proc/6553732/fd/7
<7>Use the command ln-s to create the softlink

ex:
sudo ln -s /opt/freeware/bin/pip /usr/bin/pip
