<?php
function plugin_version_officepack()
{
return array('name' => 'officepack',
'version' => '1.0',
'author'=> 'Gilles Dubois, Nicolas Derouet',
'license' => 'GPLv2',
'verMinOcs' => '2.2');
}

function plugin_init_officepack()
{
$object = new plugins;
$object -> add_cd_entry("officepack","software");
$object -> add_menu ("officepack","14000","officepack","Office Key Management","plugins");

// Officepack table creation

include 'sql/officepack.php';
include 'sql/officepack-guid-fr.php';

}

function plugin_delete_officepack()
{
$object = new plugins;
$object -> del_cd_entry("officepack");
$object -> del_menu ("officepack","14000","Office Key Management","plugins");

$object -> sql_query("DROP TABLE IF EXISTS `officepack_sku` , `officepack_lang` , `officepack_type` , `officepack_version` , `officepack`;");

}

?>
