#!/bin/bash
 
read -r -p "Are You Sure? [Y/n] " input 
case $input in
    [yY][eE][sS]|[yY])
 echo "Yes"
 ;;
    [nN][oO]|[nN])
 echo "No"
       ;;
    *)
 echo "Invalid input..."
 exit 1
 ;;
esac

echo "Do you want to continue?(yes/no)"
read input
if [ "$input" == "yes" ]
then
echo "continue"
fi
