rem this is used for checking out the latest source code from the SVN every day 
c:
cd  c:\temp
rem first cleanup the SVN repository
svn cleanup
svn checkout http://rli18.asiapacific.hpqcorp.net:8443/svn/PAF_Automation  9Auto -r HEAD  --force  --username Alter --password Alter >checkout.log