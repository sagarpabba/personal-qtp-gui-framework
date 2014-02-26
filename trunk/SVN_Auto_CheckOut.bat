rem this is used for checking out the latest source code from the SVN every day 
cd  C:\PAF_Automation
svn cleanup
svn checkout https://pdeauto06.fc.hp.com/svn/PAF_Automation_QTP . -r head  --force  --username Isabella --password Isabella >checkout.log
