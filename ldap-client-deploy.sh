# Shell script to deploy and configure opldap client for user authentication

echo -e "Installing openldap client and it's utils"
yum install authconfig authselect openldap-clients sssd oddjob-mkhomedir sssd sssd-tools -y >>/dev/null

# Deploy configurations
mv /etc/sssd/sssd.conf /etc/sssd/sssd.conf.bkp
cp sssd.conf /etc/sssd/sssd.conf

# Configure
https://kifarunix.com/configure-sssd-for-openldap-authentication-on-centos-8/


# Enable to start services on boot

systemctl enable --now oddjobd
systemctl enable --now sssd

# Start openldap client services
systemctl restart oddjobd
systemctl restart sssd
