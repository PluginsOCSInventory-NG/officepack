-- Nicolas DEROUET
-- 29/08/2012
-- officepack

-- officepack version 2.1.x (create)
CREATE TABLE IF NOT EXISTS `officepack` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `HARDWARE_ID` int(11) NOT NULL,
  `OFFICEVERSION` varchar(255) DEFAULT NULL,
  `PRODUCT` varchar(255) DEFAULT NULL,
  `PRODUCTID` varchar(255) DEFAULT NULL,
  `TYPE` int(11) DEFAULT NULL,
  `OFFICEKEY` varchar(255) DEFAULT NULL,
  `NOTE` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID`,`HARDWARE_ID`)
)  ENGINE=INNODB ;

-- officepack version 2.2.x (update)
ALTER TABLE `officepack` ADD COLUMN `GUID` varchar(255) DEFAULT NULL AFTER `OFFICEKEY`;
ALTER TABLE `officepack` ADD COLUMN `INSTALL` int(11) DEFAULT NULL AFTER `GUID`;
