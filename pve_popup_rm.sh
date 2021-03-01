#!/bin/sh

#######################################################
#
# Edits the proxmox Subscription file to make it
# think that it has a Subscription.
#
# Will disable the annoying login message about
# missing subscription.
#
# Tested on Proxmox PVE v5.2-1 and v6.0-4 
#
# The sed command will create a backup of the changed file.
# There is no guarantee that this will work for future versions.
# Use at your own risk!
#
# OneLiner:
# wget -q -O - 'https://gist.github.com/tavinus/08a63e7269e0f70d27b8fb86db596f0d/raw/' | /bin/sh
# curl -L -s 'https://gist.github.com/tavinus/08a63e7269e0f70d27b8fb86db596f0d/raw/' | /bin/sh
#
#######################################################

init_error() {
    local ret=1
    [ -z "$1" ] || printf "%s\n" "$1"
    [ -z "$2" ] || ret=$2
    exit $ret
}

# Original command
# sed -i.bak 's/NotFound/Active/g' /usr/share/perl5/PVE/API2/Subscription.pm && systemctl restart pveproxy.service

# Command to restart PVE Proxy and apply changes
PVEPXYRESTART='systemctl restart pveproxy.service'

# File/folder to be changed
TGTPATH='/usr/share/perl5/PVE/API2'
TGTFILE='Subscription.pm'

# Check dependecies
SEDBIN="$(which sed)"

[ -x "$SEDBIN" ] || init_error "Could not find 'sed' binary, aborting..."

# This will also create a .bak file with the original file contents
sed -i.bak 's/NotFound/Active/g' "$TGTPATH/$TGTFILE" && $PVEPXYRESTART

r=$?
if [ $r -eq 0 ]; then
    printf "%s\n" "All done! Please refresh your browser and test the changes!"
    exit 0
fi

printf "%s\n" "An error was detected! Changes may not have been applied!"
exit 1
