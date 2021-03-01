# Shell script to install and update newly provisioned VM's

# Update OS
echo "Updating System"
yum update -y >> /dev/null && yum upgrade -y >> /dev/null
echo "Update completed"

echo -e "Installing basic utils: \n wget \n net-tools"
yum install -y wget tree net-tools >> /dev/null

echo -e "Installing Storage utils: \n nfs \n iscsi"
yum install -y iscsi-initiator-utils nfs-utils bind-utils >> /dev/null

echo -e "Installing scm clients: \n svn \n git"
yum install -y svn subversion git >> /dev/null

echo -e "\nBasic utils installation completed"
