First remove the enterprise enterprise.proxmox repository 
# rm /etc/apt/sources.list.d/pve.enterprise.proxmox.com
#apt update && apt -y full-upgrade 
Add the no-subscription or pve-test repositori : 
#nano /etc/apt/sources.list
deb http://download.proxmox.com/debian buster pve-no-subscription
#apt update && apt -y full-upgrade 
#reboot