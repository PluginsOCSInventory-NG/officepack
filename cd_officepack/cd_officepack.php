<?php
//====================================================================================
// OCS INVENTORY REPORTS
// Copyleft Erwan GOALOU 2010 (erwan(at)ocsinventory-ng(pt)org)
// Web: http://www.ocsinventory-ng.org
//
// This code is open source and may be copied and modified as long as the source
// code is always made freely available.
// Please refer to the General Public Licence http://www.gnu.org/ or Licence.txt
//====================================================================================
 
if(AJAX) {
	parse_str($protectedPost['ocs']['0'], $params);
	$protectedPost+=$params;
	ob_start();
	$ajax = true;
} else {
	$ajax=false;
}

print_item_header($l->g(23004));

if (!isset($protectedPost['SHOW'])) {
	$protectedPost['SHOW'] = 'NOSHOW';
}

$form_name="officepack";
$table_name=$form_name;
$tab_options=$protectedPost;
$tab_options['form_name']=$form_name;
$tab_options['table_name']=$table_name;

echo open_form($form_name);

$list_fields=array( $l->g(23005) => 'PRODUCT',
$l->g(23006) => 'OFFICEVERSION',
$l->g(23007) => 'TYPE',
$l->g(23008) => 'OFFICEKEY',
);

$list_col_cant_del=$list_fields;
$default_fields= $list_fields;

$sql=prepare_sql_tab($list_fields);
$sql['SQL']  .= "FROM officepack WHERE (hardware_id = $systemid)";
array_push($sql['ARG'],$systemid);

$tab_options['ARG_SQL']=$sql['ARG'];
$tab_options['ARG_SQL_COUNT']=$systemid;

ajaxtab_entete_fixe($list_fields,$default_fields,$tab_options,$list_col_cant_del);

echo close_form();

if ($ajax) {
	ob_end_clean();
	tab_req($list_fields,$default_fields,$list_col_cant_del,$sql['SQL'],$tab_options);
	ob_start();
}
?>
