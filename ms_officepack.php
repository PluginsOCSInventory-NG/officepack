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
 
if(AJAX){
        parse_str($protectedPost['ocs']['0'], $params);
        $protectedPost+=$params;
        ob_start();
        $ajax = true;
}
else{
        $ajax=false;
}
printEnTete("Office Key Manager");
$form_name="officekey";

$data_on = array(
    "1" => "Licenses number by version",
    "2" => "Licenses by administrative data"
);

if(!isset($protectedPost['onglet'])){
    $protectedPost['onglet'] = 1;
}

$table_name=$form_name;
$tab_options=$protectedPost;
$tab_options['form_name']=$form_name;
$tab_options['table_name']=$table_name;

echo open_form($form_name);
onglet($data_on, $form_name, "onglet", 2);

if($protectedPost['onglet'] == 1){

    $sql = "SELECT OFFICEVERSION,COUNT(*) as NUMBER FROM `officepack`GROUP BY OFFICEVERSION";

    $list_fields=array(
        'Office Version' => 'OFFICEVERSION',
        'Number' => 'NUMBER',
    );

    $list_col_cant_del=$list_fields;
    $default_fields= $list_fields;

    ajaxtab_entete_fixe($list_fields,$default_fields,$tab_options,$list_col_cant_del);

}else{
    
    // select account info for sorting
    $account_info_list_sql = "Select ID, COMMENT from accountinfo_config WHERE ACCOUNT_TYPE = 'COMPUTERS'";
    $account_info_list = mysql2_query_secure($account_info_list_sql, $_SESSION['OCS']["readServer"]);
    
    echo "<p>Accountinfo : <select name='accountinfo' onchange='this.form.submit();'>";
    while ($row = mysqli_fetch_object($account_info_list)) {
        $id = $row->ID;
        $str = $row->COMMENT;
        if(isset($protectedPost['accountinfo']) && $protectedPost['accountinfo'] == $row->ID){
            echo "<option value='$id' selected>$str</option> ";
        }else{
            echo "<option value='$id'>$str</option> ";
        }
    }
    echo "</select></p>";
    
    // Select which office version we want to see
    $sql_office = "SELECT OFFICEVERSION FROM `officepack`GROUP BY OFFICEVERSION";
    $result = mysql2_query_secure($sql_office, $_SESSION['OCS']["readServer"]);
  
    echo "<p>Office version : <select name='officeversion' onchange='this.form.submit();'>";
    while ($row = mysqli_fetch_object($result)) {
        $officeversion = $row->OFFICEVERSION;
        if(isset($protectedPost['officeversion']) && $protectedPost['officeversion'] == $row->OFFICEVERSION){
            echo "<option value='$officeversion' selected>$officeversion</option> ";
        }else{
            echo "<option value='$officeversion'>$officeversion</option> ";
        }
    }
    echo "</select></p>";
    
    if( isset($protectedPost['accountinfo']) && isset($protectedPost['officeversion'])){
        $fields = "fields_".$protectedPost['accountinfo'];
        if($protectedPost['accountinfo']){
            $fields = "TAG";
        }
        $office = $protectedPost['officeversion'];

        $sql = "SELECT a.".$fields." as ACC , COUNT(".$fields.") as ACCNB FROM `accountinfo` as a INNER JOIN officepack as o ON a.hardware_id = o.hardware_id WHERE o.officeversion = '".$office."' GROUP BY ".$fields."";
        
        $list_fields=array(
            'Accountinfo' => "ACC",
            'Licenses number' => 'ACCNB',
        );

        $list_col_cant_del=$list_fields;
        $default_fields= $list_fields;

        ajaxtab_entete_fixe($list_fields,$default_fields,$tab_options,$list_col_cant_del);
    }
    
    
}

echo close_form();
if ($ajax){
        ob_end_clean();
        tab_req($list_fields,$default_fields,$list_col_cant_del,$sql,$tab_options);
        ob_start();
}

function getAccountinfoNumber($id){
    
}

?>
