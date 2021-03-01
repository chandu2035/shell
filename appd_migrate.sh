#!/bin/bash

###########################################################################
## This scirpt is to migrate AppD on-prem controller to point AppD SaaS  ##
###########################################################################

################ Update UAT AppD SaaS Controller Values ###################
UAT_SAAS_C_HOST=uat.appd.saas.host
UAT_SAAS_C_PORT=8443
UAT_SAAS_C_SSL=true
UAT_SAAS_C_ACCOUNT=uat_env
UAT_SAAS_C_KEY=d5734ksdd8CSRUAT
UAT_APPD_PROXYHOST=172.0.0.128
UAT_APPD_PROXYPORT=80

################ Update PROD AppD SaaS Controller Values ##################
PROD_SAAS_C_HOST=prod.saas.appd
PROD_SAAS_C_PORT=443
PROD_SAAS_C_SSL=true
PROD_SAAS_C_ACCOUNT=prod_env
PROD_SAAS_C_KEY=prd12345intePROD
PROD_APPD_PROXYHOST=10.163.0.100
PROD_APPD_PROXYPORT=8080

###########################################################################
APPD_HOME_DIR=$1
APP_AGENT_STARTUP_FILE=$2
ENVIRONMENT=$3
CONTROLLER_CONFIG_FILE=$APPD_HOME_DIR/ver/conf/controller-info.xml

###########################################################################
display_usage () {
   echo -e "\nPlease pass below required valid arguments:"
   echo -e "Usage: $0 [appd_base_dir] [/path/startup file] [environment: uat/prod]\n"
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

check_perms () {
   if [ -w "$CONTROLLER_CONFIG_FILE" ]; then
   echo -e "User has write permissions: $CONTROLLER_CONFIG_FILE"
   else
   echo -e "User doesn't have write permissions on $CONTROLLER_CONFIG_FILE"
   exit 1
   fi
   if [ -w "$APP_AGENT_STARTUP_FILE" ]; then
   echo -e "User has write permissions: $APP_AGENT_STARTUP_FILE"
   else
   echo -e "User doesn't have write permissions on $APP_AGENT_STARTUP_FILE"
   exit 1
   fi
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

########## Extract and print required parameters ############################
app_startup_config () {
   grep "^[^#;]" $APP_AGENT_STARTUP_FILE | grep  -e 'javaagent.jar'
}

########## Extract and print required parameters ############################
print_current_conf () {
   echo -e "\nCurrent configurations of $CONTROLLER_CONFIG_FILE"
   echo -e "=========================================================="
   HOST=$(sed -n 's:.*<controller-host>\(.*\)</controller-host>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   PORT=$(sed -n 's:.*<controller-port>\(.*\)</controller-port>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   SSL=$(sed -n 's:.*<controller-ssl-enabled>\(.*\)</controller-ssl-enabled>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   ACCOUNT=$(sed -n 's:.*<account-name>\(.*\)</account-name>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   KEY=$(sed -n 's:.*<account-access-key>\(.*\)</account-access-key>.*:\1:p' $CONTROLLER_CONFIG_FILE)
   echo -e "Controller Host: $HOST\nController Port: $PORT\nSSL Enabled: $SSL\nAccount Name: $ACCOUNT\nAccess Key: $KEY"
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

########## UAT Update APPD SaaS Controller ##################################
update_uat_controller_conf () {
   echo -e "\nUPDATED APPD SaaS CONFIGURATION in $CONTROLLER_CONFIG_FILE"
   echo -e "=========================================================="
   sed -i '' -e "s/${HOST}/${UAT_SAAS_C_HOST}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${PORT}/${UAT_SAAS_C_PORT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${SSL}/${UAT_SAAS_C_SSL}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${ACCOUNT}/${UAT_SAAS_C_ACCOUNT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${KEY}/${UAT_SAAS_C_KEY}/g" $CONTROLLER_CONFIG_FILE
}

########## Update UAT APPD SaaS Proxy parameters ############################
function update_uat_app_startup_config () {
   CHECK_PROXY=$(grep -e '-Dappdynamics.http.proxy' $APP_AGENT_STARTUP_FILE)
   if [ $? != 0 ]; then
       # Update AppD Proxy Values below
       #sed -i '' -e '/^#/!s/javaagent\.jar.*/& -Dappdynamics.http.proxyHost='"$UAT_APPD_PROXYHOST"'/' $APP_AGENT_STARTUP_FILE
       sed -i 's#javaagent.jar#javaagent.jar -Dappdynamics.http.proxyHost='"$UAT_APPD_PROXYHOST"' -Dappdynamics.http.proxyPort='"$UAT_APPD_PROXYPORT"''
       #sed -i '' -e '/^#/!s/javaagent\.jar.*/& -Dappdynamics.http.proxyPort='"$UAT_APPD_PROXYPORT"'/' $APP_AGENT_STARTUP_FILE
       echo -e "\nProxy parameters appended:"
       else
       echo -e "\nProxy parameters already available"
   fi
}

########## UAT Update APPD SaaS Controller ##################################
update_prod_controller_conf () {
   echo -e "\nUPDATED APPD SaaS CONFIGURATION in $CONTROLLER_CONFIG_FILE"
   echo -e "=========================================================="
   sed -i '' -e "s/${HOST}/${PROD_SAAS_C_HOST}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${PORT}/${PROD_SAAS_C_PORT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${SSL}/${PROD_SAAS_C_SSL}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${ACCOUNT}/${PROD_SAAS_C_ACCOUNT}/g" $CONTROLLER_CONFIG_FILE
   sed -i '' -e "s/${KEY}/${PROD_SAAS_C_KEY}/g" $CONTROLLER_CONFIG_FILE
}

########## Update PROD APPD SaaS Proxy parameters ###########################
update_prod_app_startup_config () {
   CHECK_PROXY=$(grep -e '-Dappdynamics.http.proxy' $APP_AGENT_STARTUP_FILE)
   if [ $? != 0 ]; then
       # Update AppD Proxy Values below
       sed -i '' -e '/^#/!s/javaagent\.jar.*/& -Dappdynamics.http.proxyHost='"$PROD_APPD_PROXYHOST"'/' $APP_AGENT_STARTUP_FILE
       sed -i '' -e '/^#/!s/javaagent\.jar.*/& -Dappdynamics.http.proxyPort='"$PROD_APPD_PROXYPORT"'/' $APP_AGENT_STARTUP_FILE
       echo -e "\nProxy parameters appended:"
       else
       echo -e "\nProxy parameters already available:"
   fi
}

########## UAT Migration ######################################################
uat_migration () {
   echo -e "\nAppD migration prerequisites check"
   echo -e "=========================================================="
   check_appd_dir
   check_app_startup_file
   check_perms
   print_current_conf
   echo -e "\nPresent AppD parameters:"
   app_startup_config
   confirm_execution
   backup_conf_files
   update_uat_controller_conf
   print_current_conf
   update_uat_app_startup_config
   app_startup_config
}

########## PROD Migration ######################################################
prod_migration () {
   echo -e "\nAppD migration prerequisites check"
   echo -e "=========================================================="
   check_appd_dir
   check_app_startup_file
   check_perms
   print_current_conf
   echo -e "\nPresent AppD parameters:"
   app_startup_config
   confirm_execution
   backup_conf_files
   update_prod_controller_conf
   print_current_conf
   update_prod_app_startup_config
   app_startup_config
}

################################################################################
if [ $# != 3 ]; then
   display_usage
   exit 1
fi

if [[ $3 = "UAT" || $3 = "uat" ]]; then
   echo -e "\nUAT migration:"
   echo -e "=========================================================="
   uat_migration
   elif [[ $3 = "PROD" || $3 = "prod" ]]; then
   echo -e "\nPROD migration:"
   echo -e "=========================================================="
   prod_migration
   else
   echo -e "\nInvalid entry!"
   display_usage
   exit 1
fi

###############################################################################
##                         END OF THE SCRIPT                                 ##
###############################################################################