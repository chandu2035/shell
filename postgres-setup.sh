
sudo yum install -y https://download.postgresql.org/pub/repos/yum/reporpms/EL-8-x86_64/pgdg-redhat-repo-latest.noarch.rpm
# rpm -qi pgdg-redhat-repo
sudo dnf -qy module disable postgresql
sudo dnf -y install postgresql12 postgresql12-server
sudo  /usr/pgsql-12/bin/postgresql-12-setup initdb
sudo systemctl enable postgresql-12
sudo systemctl start postgresql-12

sudo firewall-cmd --add-service=postgresql --permanent
sudo firewall-cmd --reload
sudo su - postgres
psql -c "alter user postgres with password 'admin123'"


sudo dnf -y install https://dl.fedoraproject.org/pub/epel/epel-release-latest-8.noarch.rpm
sudo dnf config-manager --set-enabled PowerTools
sudo dnf -y install https://download.postgresql.org/pub/repos/yum/reporpms/EL-8-x86_64/pgdg-redhat-repo-latest.noarch.rpm
sudo dnf -qy module disable postgresql
sudo dnf -y install pgadmin4
sudo systemctl start httpd && sudo systemctl enable httpd
sudo cp /etc/httpd/conf.d/pgadmin4.conf.sample /etc/httpd/conf.d/pgadmin4.conf
sudo httpd -t
sudo systemctl restart httpd

sudo mkdir -p /var/lib/pgadmin4/ /var/log/pgadmin4/
sudo dnf -y install python3-bcrypt python3-pynacl

sudo chown -R apache:apache /var/lib/pgadmin4 /var/log/pgadmin4
sudo chown -R apache:apache /var/lib/pgadmin4 /var/log/pgadmin4

sudo semanage fcontext -a -t httpd_sys_rw_content_t "/var/lib/pgadmin4(/.*)?"
sudo semanage fcontext -a -t httpd_sys_rw_content_t "/var/log/pgadmin4(/.*)?"
sudo restorecon -Rv /var/lib/pgadmin4/
sudo restorecon -Rv /var/log/pgadmin4/


sudo systemctl restart httpd
sudo firewall-cmd --permanent --add-service=http
sudo firewall-cmd --reload
