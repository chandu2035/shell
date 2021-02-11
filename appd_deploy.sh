#!/bin/bash

###########################################################################
## This scirpt is to migrate AppD on-prem controller to point AppD SaaS  ##
###########################################################################

###########################################################################
APPD_HOME_DIR=$1
APP_AGENT_STARTUP_FILE=$2
CONTROLLER_CONFIG_FILE=$APPD_HOME_DIR/ver/conf/controller-info.xml

################ Update AppD SaaS Controller Values #######################
SAAS_C_HOST=192.168.0.101
SAAS_C_PORT=443
SAAS_C_SSL=false
SAAS_C_ACCOUNT=testenv
SAAS_C_KEY=d5734ksdd8CSR

########## Update APPD SaaS Proxy parameters ################################
update_app_startup_config () {
   CHECK_PROXY=$(grep -e '-Dappdynamics.http.proxy' $APP_AGENT_STARTUP_FILE)
   if [ $? != 0 ]; then
       # Update AppD Proxy Values below
       sed -i ‘’ -e '/-Dappdynamics.agent.nodename/ s/$/ -Dappdynamics.http.proxyHost=10.XXX.XXX.XXX/' $APP_AGENT_STARTUP_FILE
       sed -i ‘’ -e '/-Dappdynamics.agent.nodename/ s/$/ -Dappdynamics.http.proxyPort=80/' $APP_AGENT_STARTUP_FILE
       echo -e "\nProxy parameters appended:"
       else
       echo -e "\nProxy parameters already available"
   fi
}

############### Prerequisites check and etc.##################################
display_usage () {
   echo -e "\nPlease pass 2 arguments"
   echo -e "Usage: $0 [AppDynamics_Base_Dir] [Application_startup_file_with_absolute_path]\n"
}

confirm_execution () {
   echo -e "\n"
   read -r -p "Would you like to proceed ? [Y/n] " input 
   case $input in
      yY][eE][sS]|[yY])
      echo "Proceeding with changes......."
      ;;
      [nN][oO]|[nN])
      echo "Cancelled execution."
      exit
      ;;
      *)
      echo "Invalid input..."
      exit 1
      ;;
   esac
}

check_appd_dir () {
   if [ -d "$APPD_HOME_DIR" ];  then
   echo -e "AppDynamics Base directory: $APPD_HOME_DIR exists"
   else
   echo -e "$APPD_HOME_DIR Directory NOT found"
   exit 1
   fi
}

check_app_startup_file () {
   if [ -f "$APP_AGENT_STARTUP_FILE" ]; then
   echo -e "File: $APP_AGENT_STARTUP_FILE exists"
   else
   echo -e "$APP_AGENT_STARTUP_FILE File DOES NOT exists"
   exit 1
   fi
}
########## Backup config files ##############################################
backup_conf_files () {
   echo -e "\nBackup Config FIles"
   echo -e "=========================================================="
   TIMESTAMP=$(date +%Y%m%d_%H%M%S)
   cp $CONTROLLER_CONFIG_FILE $CONTROLLER_CONFIG_FILE.bkp_$TIMESTAMP
   cp $APP_AGENT_STARTUP_FILE $APP_AGENT_STARTUP_FILE.bkp_$TIMESTAMP
   echo -e "Backup has been completed for below files:\n$CONTROLLER_CONFIG_FILE\n$APP_AGENT_STARTUP_FILE\n"
}
########## Extract and print required parameters ############################
print_current_conf () {
   HOST=$(sed -n 's:.*<controller-host>\(.*\)</controller-host>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   PORT=$(sed -n 's:.*<controller-port>\(.*\)</controller-port>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   SSL=$(sed -n 's:.*<controller-ssl-enabled>\(.*\)</controller-ssl-enabled>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   ACCOUNT=$(sed -n 's:.*<account-name>\(.*\)</account-name>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   KEY=$(sed -n 's:.*<account-access-key>\(.*\)</account-access-key>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   echo -e "Controller Host: $HOST\nController Port: $PORT\nSSL Enabled: $SSL\nAccount Name: $ACCOUNT\nAccess Key: $KEY"
}
########## Update APPD SaaS parameters ######################################
update_controller_conf () {
   sed -i '' -e "s/${HOST}/${SAAS_C_HOST}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${PORT}/${SAAS_C_PORT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${SSL}/${SAAS_C_SSL}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${ACCOUNT}/${SAAS_C_ACCOUNT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${KEY}/${SAAS_C_KEY}/g" $CONTROLLER_CONFIG_FILE
}
########## Extract and print required parameters ############################
app_startup_config () {
   grep -e '-Dappdynamics.agent' $APP_AGENT_STARTUP_FILE
}

if [ $# != 2 ]; then
   display_usage
   exit 1
fi
########## Main fuction ####################################################
main () {
   echo -e "\nPrerequisites check"
   echo -e "=========================================================="
   check_appd_dir
   check_app_startup_file
   echo -e "\nCurrent configurations of $CONTROLLER_CONFIG_FILE"
   echo -e "=========================================================="
   print_current_conf
   echo -e "\nPresent AppD parameters:"
   app_startup_config
   confirm_execution
   backup_conf_files
   update_controller_conf
   echo -e "\nUPDATED APPD SaaS CONFIGURATION in $CONTROLLER_CONFIG_FILE"
   echo -e "=========================================================="
   print_current_conf
   update_app_startup_config
   app_startup_config
}

########## Execute main funtion to complete the job ########################
main

############################################################################
##                      END OF THE SCRIPT                                 ##
############################################################################