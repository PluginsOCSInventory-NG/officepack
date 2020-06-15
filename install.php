<?php

/**
 * This function is called on installation and is used to create database schema for the plugin
 */
function extension_install_officepack()
{
    $commonObject = new ExtensionCommon;

    // Remove older tables
    $commonObject -> sqlQuery("DROP TABLE IF EXISTS `officepack_sku` , `officepack_lang` , `officepack_type` , `officepack_version` , `officepack`;");

    // Install tables
    $commonObject -> sqlQuery("CREATE TABLE IF NOT EXISTS `officepack` (
        `ID` int(11) NOT NULL AUTO_INCREMENT,
        `HARDWARE_ID` int(11) NOT NULL,
        `OFFICEVERSION` varchar(255) DEFAULT NULL,
        `PRODUCT` varchar(255) DEFAULT NULL,
        `PRODUCTID` varchar(255) DEFAULT NULL,
        `TYPE` int(11) DEFAULT NULL,
        `OFFICEKEY` varchar(255) DEFAULT NULL,
        `GUID` varchar(255) DEFAULT NULL,
        `INSTALL` int(11) DEFAULT NULL,
        `NOTE` varchar(255) DEFAULT NULL,
        PRIMARY KEY (`ID`,`HARDWARE_ID`)
      )  ENGINE=INNODB ;");
      
}

/**
 * This function is called on removal and is used to destroy database schema for the plugin
 */
function extension_delete_officepack()
{
    $commonObject = new ExtensionCommon;
    $commonObject -> sqlQuery("DROP TABLE IF EXISTS `officepack_sku` , `officepack_lang` , `officepack_type` , `officepack_version` , `officepack`;");
}

/**
 * This function is called on plugin upgrade
 */
function extension_upgrade_officepack()
{

}
