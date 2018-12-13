<?php

/**
 * This function is called on installation and is used to create database schema for the plugin
 */
function extension_install_officepack()
{
    $commonObject = new ExtensionCommon;

    // Officepack table creation

    include 'sql/officepack.php';
    include 'sql/officepack-guid-fr.php';
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
