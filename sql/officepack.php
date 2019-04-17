<?php

// Create table
$commonObject -> sqlQuery("CREATE TABLE IF NOT EXISTS `officepack` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `HARDWARE_ID` int(11) NOT NULL,
  `OFFICEVERSION` varchar(255) DEFAULT NULL,
  `PRODUCT` varchar(255) DEFAULT NULL,
  `PRODUCTID` varchar(255) DEFAULT NULL,
  `TYPE` int(11) DEFAULT NULL,
  `OFFICEKEY` varchar(255) DEFAULT NULL,
  `NOTE` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID`,`HARDWARE_ID`)
)  ENGINE=INNODB ;");

//Alter data table
$commonObject -> sqlQuery("
		ALTER TABLE `officepack` ADD COLUMN `GUID` varchar(255) DEFAULT NULL AFTER `OFFICEKEY`;
		ALTER TABLE `officepack` ADD COLUMN `INSTALL` int(11) DEFAULT NULL AFTER `GUID`;");

?>
