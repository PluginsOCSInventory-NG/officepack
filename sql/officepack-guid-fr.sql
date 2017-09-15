-- Nicolas DEROUET
-- 14/01/2013 17:30
-- officepack-guid (Français)

--
-- Structure de la table `officepack_sku`
--
DROP TABLE IF EXISTS `officepack_sku`;
CREATE TABLE IF NOT EXISTS `officepack_sku` (
  `VERSION` varchar(109) DEFAULT NULL,
  `REF_ID` varchar(127) DEFAULT NULL,
  `PRODUCT` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`VERSION`,`REF_ID`)
) ENGINE=INNODB;


-- Office 2000 : http://support.microsoft.com/kb/230848  
INSERT INTO `officepack_sku` VALUES ('2000','00','Microsoft Office 2000 Premium Edition CD1');
INSERT INTO `officepack_sku` VALUES ('2000','01','Microsoft Office 2000 Professional Edition');
INSERT INTO `officepack_sku` VALUES ('2000','02','Microsoft Office 2000 Standard Edition');
INSERT INTO `officepack_sku` VALUES ('2000','03','Microsoft Office 2000 Small Business Edition');
INSERT INTO `officepack_sku` VALUES ('2000','04','Microsoft Office 2000 Premium CD2');
INSERT INTO `officepack_sku` VALUES ('2000','05','Office CD2 SMALL');
INSERT INTO `officepack_sku` VALUES ('2000','10','Microsoft Access 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','11','Microsoft Excel 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','12','Microsoft FrontPage 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','13','Microsoft PowerPoint 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','14','Microsoft Publisher 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','15','Office Server Extensions');
INSERT INTO `officepack_sku` VALUES ('2000','16','Microsoft Outlook 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','17','Microsoft Word 2000 (standalone)');
INSERT INTO `officepack_sku` VALUES ('2000','18','Microsoft Access 2000 runtime version');
INSERT INTO `officepack_sku` VALUES ('2000','19','FrontPage Server Extensions');
INSERT INTO `officepack_sku` VALUES ('2000','1A','Publisher Standalone OEM');
INSERT INTO `officepack_sku` VALUES ('2000','1B','DMMWeb');
INSERT INTO `officepack_sku` VALUES ('2000','1C','FP WECCOM');
INSERT INTO `officepack_sku` VALUES ('2000','40','Publisher Trial CD');
INSERT INTO `officepack_sku` VALUES ('2000','41','Publisher Trial Web');
INSERT INTO `officepack_sku` VALUES ('2000','42','SBB');
INSERT INTO `officepack_sku` VALUES ('2000','43','SBT');
INSERT INTO `officepack_sku` VALUES ('2000','44','SBT CD2');
INSERT INTO `officepack_sku` VALUES ('2000','45','SBTART');
INSERT INTO `officepack_sku` VALUES ('2000','46','Web Components');
INSERT INTO `officepack_sku` VALUES ('2000','47','VP Office CD2 with LVP');
INSERT INTO `officepack_sku` VALUES ('2000','48','VP PUB with LVP');
INSERT INTO `officepack_sku` VALUES ('2000','49','VP PUB with LVP OEM');
INSERT INTO `officepack_sku` VALUES ('2000','4F','Access 2000 SR-1 Run-Time Minimum');

-- Office XP : http://support.microsoft.com/kb/302663
INSERT INTO `officepack_sku` VALUES ('XP','11','Microsoft Office XP Edition Professionnelle');
INSERT INTO `officepack_sku` VALUES ('XP','12','Microsoft Office XP Edition Standard ');
INSERT INTO `officepack_sku` VALUES ('XP','13','Microsoft Office XP Édition PME');
INSERT INTO `officepack_sku` VALUES ('XP','14','Serveur Web Microsoft Office XP');
INSERT INTO `officepack_sku` VALUES ('XP','15','Microsoft Access 2002');
INSERT INTO `officepack_sku` VALUES ('XP','16','Microsoft Excel 2002');
INSERT INTO `officepack_sku` VALUES ('XP','17','Microsoft FrontPage 2002');
INSERT INTO `officepack_sku` VALUES ('XP','18','Microsoft PowerPoint 2002');
INSERT INTO `officepack_sku` VALUES ('XP','19','Microsoft Publisher 2002');
INSERT INTO `officepack_sku` VALUES ('XP','1A','Microsoft Outlook 2002');
INSERT INTO `officepack_sku` VALUES ('XP','1B','Microsoft Word 2002');
INSERT INTO `officepack_sku` VALUES ('XP','1C','Microsoft Access 2002 Runtime');
INSERT INTO `officepack_sku` VALUES ('XP','1D','Extensions serveur 2002 Microsoft FrontPage');
INSERT INTO `officepack_sku` VALUES ('XP','1E','Pack de l''interface utilisateur multilingue Microsoft Office');
INSERT INTO `officepack_sku` VALUES ('XP','1F','Kit d''outils de vérification orthographique Microsoft Office');
INSERT INTO `officepack_sku` VALUES ('XP','20','Mise à jour des fichiers systèmes');
INSERT INTO `officepack_sku` VALUES ('XP','22','non utilisé');
INSERT INTO `officepack_sku` VALUES ('XP','23','Assistant Pack de l''interface utilisateur multilingue Microsoft Office');
INSERT INTO `officepack_sku` VALUES ('XP','24','Kit de ressources Microsoft Office XP');
INSERT INTO `officepack_sku` VALUES ('XP','25','Outils du kit de ressources Microsoft Office XP (téléchargement à partir du Web)');
INSERT INTO `officepack_sku` VALUES ('XP','26','Composants Web Microsoft Office');
INSERT INTO `officepack_sku` VALUES ('XP','27','Microsoft Project 2002');
INSERT INTO `officepack_sku` VALUES ('XP','28','Microsoft Office XP Professionnel avec FrontPage');
INSERT INTO `officepack_sku` VALUES ('XP','29','Abonnement Microsoft Office XP Edition Professionnelle');
INSERT INTO `officepack_sku` VALUES ('XP','2A','Abonnement Microsoft Office XP Edition PME');
INSERT INTO `officepack_sku` VALUES ('XP','2B','Microsoft Publisher 2002 Deluxe Edition');
INSERT INTO `officepack_sku` VALUES ('XP','2F','IME autonome (JPN uniquement)');
INSERT INTO `officepack_sku` VALUES ('XP','30','Contenu Microsoft Office XP Media');
INSERT INTO `officepack_sku` VALUES ('XP','31','Client Web Microsoft Project 2002');
INSERT INTO `officepack_sku` VALUES ('XP','32','Serveur Web Microsoft Project 2002');
INSERT INTO `officepack_sku` VALUES ('XP','33','Microsoft Office XP PIPC1 (PC pré-installé) (JPN uniquement)');
INSERT INTO `officepack_sku` VALUES ('XP','34','Microsoft Office XP PIPC2 (PC pré-installé) (JPN uniquement)');
INSERT INTO `officepack_sku` VALUES ('XP','35','Contenu de luxe Microsoft Office XP Media');
INSERT INTO `officepack_sku` VALUES ('XP','3A','Project 2002 Standard');
INSERT INTO `officepack_sku` VALUES ('XP','3B','Project 2002 Professional');
INSERT INTO `officepack_sku` VALUES ('XP','51','Microsoft Office Visio Professionnel 2002');
INSERT INTO `officepack_sku` VALUES ('XP','54','Microsoft Office Visio Standard 2002');

-- Office 2003 : http://support.microsoft.com/kb/832672
INSERT INTO `officepack_sku` VALUES ('2003','11','Microsoft Office Édition Professionnelle Entreprise 2003');
INSERT INTO `officepack_sku` VALUES ('2003','12','Microsoft Office Édition Standard 2003');
INSERT INTO `officepack_sku` VALUES ('2003','13','Microsoft Office Édition Basique 2003');
INSERT INTO `officepack_sku` VALUES ('2003','14','Windows Windows SharePoint Services 2.0');
INSERT INTO `officepack_sku` VALUES ('2003','15','Microsoft Office Access 2003');
INSERT INTO `officepack_sku` VALUES ('2003','16','Microsoft Office Excel 2003');
INSERT INTO `officepack_sku` VALUES ('2003','17','Microsoft Office FrontPage 2003');
INSERT INTO `officepack_sku` VALUES ('2003','18','Microsoft Office PowerPoint 2003');
INSERT INTO `officepack_sku` VALUES ('2003','19','Microsoft Office Publisher 2003');
INSERT INTO `officepack_sku` VALUES ('2003','1A','Microsoft Office Outlook Professionnel 2003');
INSERT INTO `officepack_sku` VALUES ('2003','1B','Microsoft Office Word 2003');
INSERT INTO `officepack_sku` VALUES ('2003','1C','Microsoft Office Access 2003 Runtime');
INSERT INTO `officepack_sku` VALUES ('2003','1E','Pack d''interface utilisateur Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','1F','Outils de vérification linguistique Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','23','Pack d''interface utilisateur multilingue Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','24','Kit de ressources Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','26','Composants Web Microsoft Office XP');
INSERT INTO `officepack_sku` VALUES ('2003','2E','Microsoft Office 2003 Research Service SDK');
INSERT INTO `officepack_sku` VALUES ('2003','44','Microsoft Office InfoPath 2003');
INSERT INTO `officepack_sku` VALUES ('2003','83','Visionneuse HTML Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','92','Windows SharePoint Services 2.0 Lot de modèles en anglais');
INSERT INTO `officepack_sku` VALUES ('2003','93','Microsoft Office 2003 Web Parts and Components en anglais');
INSERT INTO `officepack_sku` VALUES ('2003','A1','Microsoft Office OneNote 2003');
INSERT INTO `officepack_sku` VALUES ('2003','A4','Composants Web Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','A5','Outil de migration Microsoft SharePoint 2003');
INSERT INTO `officepack_sku` VALUES ('2003','AA','Diffusion de présentation Microsoft Office PowerPoint 2003');
INSERT INTO `officepack_sku` VALUES ('2003','AB','Microsoft Office PowerPoint 2003 Lot de modèles 1');
INSERT INTO `officepack_sku` VALUES ('2003','AC','Microsoft Office PowerPoint 2003 Lot de modèles 2');
INSERT INTO `officepack_sku` VALUES ('2003','AD','Microsoft Office PowerPoint 2003 Lot de modèles 3');
INSERT INTO `officepack_sku` VALUES ('2003','AE','Organigramme hiérarchique Microsoft 2.0');
INSERT INTO `officepack_sku` VALUES ('2003','CA','Microsoft Office Édition PME 2003');
INSERT INTO `officepack_sku` VALUES ('2003','D0','Microsoft Office Access 2003 Developer Extensions');
INSERT INTO `officepack_sku` VALUES ('2003','DC','SDK Microsoft Office 2003 Smart Document');
INSERT INTO `officepack_sku` VALUES ('2003','E0','Microsoft Office Outlook Standard 2003');
INSERT INTO `officepack_sku` VALUES ('2003','E3','Microsoft Office Édition Professionnelle 2003 (avec InfoPath 2003)');
INSERT INTO `officepack_sku` VALUES ('2003','FD','Microsoft Office Outlook 2003 (distribué par MSN)');
INSERT INTO `officepack_sku` VALUES ('2003','FF','Pack linguistique LIP de Microsoft Office 2003');
INSERT INTO `officepack_sku` VALUES ('2003','F8','Outil de suppression des métadonnées');
INSERT INTO `officepack_sku` VALUES ('2003','3A','Microsoft Office Project Standard 2003');
INSERT INTO `officepack_sku` VALUES ('2003','3B','Microsoft Office Project Professionnel 2003');
INSERT INTO `officepack_sku` VALUES ('2003','32','Microsoft Office Project Server 2003');
INSERT INTO `officepack_sku` VALUES ('2003','51','Microsoft Office Visio Professionnel 2003');
INSERT INTO `officepack_sku` VALUES ('2003','53','Microsoft Office Visio Standard 2003');
INSERT INTO `officepack_sku` VALUES ('2003','5E','Pack d''interface utilisateur multilingue Microsoft Office Visio 2003');

-- Office 2003 : http://support.microsoft.com/kb/832672?ln=en-en
INSERT INTO `officepack_sku` VALUES ('2003','52','Microsoft Office Visio Viewer 2003');
INSERT INTO `officepack_sku` VALUES ('2003','55','Microsoft Office Visio pour Enterprise Architects 2003');

-- Office 2007 : http://support.microsoft.com/kb/928516
INSERT INTO `officepack_sku` VALUES ('2007','0011','Microsoft Office Professionnel Plus 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0012','Microsoft Office Standard 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0013','Microsoft Office Basic 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0014','Microsoft Office Professionnel 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0015','Microsoft Office Access 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0016','Microsoft Office Excel 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0017','Microsoft Office SharePoint Designer 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0018','Microsoft Office PowerPoint 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0019','Microsoft Office Publisher 2007');
INSERT INTO `officepack_sku` VALUES ('2007','001A','Microsoft Office Outlook 2007');
INSERT INTO `officepack_sku` VALUES ('2007','001B','Microsoft Office Word 2007');
INSERT INTO `officepack_sku` VALUES ('2007','001C','Microsoft Office Access Runtime 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0020','Pack de compatibilité pour formats de fichier Microsoft Office pour Word, Excel et PowerPoint 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0026','Microsoft Expression Web');
INSERT INTO `officepack_sku` VALUES ('2007','002E','Microsoft Office Ultimate 2007');
INSERT INTO `officepack_sku` VALUES ('2007','002F','Microsoft Office Édition Familial et Étudiants 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0030','Microsoft Office Édition Enterprise 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0031','Microsoft Office Professional Hybrid 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0033','Microsoft Office Personal 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0035','Microsoft Office Professional Hybrid 2007');
INSERT INTO `officepack_sku` VALUES ('2007','003A','Microsoft Office Project Standard 2007');
INSERT INTO `officepack_sku` VALUES ('2007','003B','Microsoft Office Project Professionnel 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0044','Microsoft Office InfoPath 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0051','Microsoft Office Visio Professionnel 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0052','Microsoft Office Visio Viewer 2007');
INSERT INTO `officepack_sku` VALUES ('2007','0053','Microsoft Office Visio Standard 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00A1','Microsoft Office OneNote 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00A3','Microsoft Office OneNote Home Student 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00A7','Assistant Impression de calendriers pour Microsoft Office Outlook 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00A9','Microsoft Office InterConnect 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00AF','Visionneuse Microsoft Office PowerPoint 2007 (Anglais)');
INSERT INTO `officepack_sku` VALUES ('2007','00B0','Macro complémentaire Microsoft Enregistrer en tant que PDF');
INSERT INTO `officepack_sku` VALUES ('2007','00B1','Macro complémentaire Microsoft Enregistrer en tant que XPS');
INSERT INTO `officepack_sku` VALUES ('2007','00B2','Macro complémentaire Microsoft Enregistrer en tant que PDF ou XPS');
INSERT INTO `officepack_sku` VALUES ('2007','00BA','Microsoft Office Groove 2007');
INSERT INTO `officepack_sku` VALUES ('2007','00CA','Microsoft Office Édition PME 2007');
INSERT INTO `officepack_sku` VALUES ('2007','10D7','Microsoft Office InfoPath Forms Services');
INSERT INTO `officepack_sku` VALUES ('2007','110D','Microsoft Office SharePoint Server 2007');
INSERT INTO `officepack_sku` VALUES ('2007','1122','Windows SharePoint Services Developer Resources 1.2');
INSERT INTO `officepack_sku` VALUES ('2007','0010','SKU - Mise à jour logicielle pour les dossiers Web (anglais) 12');

-- Office 2007 : no source
INSERT INTO `officepack_sku` VALUES ('2007','0021','Microsoft Office Visual Web Developer 2007');

-- Office 2010 : http://support.microsoft.com/kb/2186281 
INSERT INTO `officepack_sku` VALUES ('2010','0011','Microsoft Office Professionnel Plus 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0012','Microsoft Office Standard 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0013','Microsoft Office Famille et Petite Entreprise 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0014','Microsoft Office Professionnel 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0015','Microsoft Access 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0016','Microsoft Excel 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0017','Microsoft SharePoint Designer 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0018','Microsoft PowerPoint 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0019','Microsoft Publisher 2010');
INSERT INTO `officepack_sku` VALUES ('2010','001A','Microsoft Outlook 2010');
INSERT INTO `officepack_sku` VALUES ('2010','001B','Microsoft Word 2010');
INSERT INTO `officepack_sku` VALUES ('2010','001C','Microsoft Access Runtime 2010');
INSERT INTO `officepack_sku` VALUES ('2010','001F','Microsoft Office Proofing Tools Kit Compilation 2010');
INSERT INTO `officepack_sku` VALUES ('2010','002F','Microsoft Office Famille et Étudiant 2010');
INSERT INTO `officepack_sku` VALUES ('2010','003A','Microsoft Project Standard 2010');
INSERT INTO `officepack_sku` VALUES ('2010','003B','Microsoft Project Professionnel 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0044','Microsoft InfoPath 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0052','Visionneuse Microsoft Visio 2010');
INSERT INTO `officepack_sku` VALUES ('2010','0057','Microsoft Visio 2010');
INSERT INTO `officepack_sku` VALUES ('2010','007A','Microsoft Outlook Connector');
INSERT INTO `officepack_sku` VALUES ('2010','008B','Notions de base sur Microsoft Office PME 2010');
INSERT INTO `officepack_sku` VALUES ('2010','00A1','Microsoft OneNote 2010');
INSERT INTO `officepack_sku` VALUES ('2010','00AF','Visionneuse Microsoft PowerPoint 2010');
INSERT INTO `officepack_sku` VALUES ('2010','00BA','Microsoft Office SharePoint Workspace 2010');
INSERT INTO `officepack_sku` VALUES ('2010','110D','Microsoft Office SharePoint Server 2010');
INSERT INTO `officepack_sku` VALUES ('2010','110F','Microsoft Project Server 2010');

-- Office 2010 : no source
INSERT INTO `officepack_sku` VALUES ('2010','003D','Microsoft Office Single Image 2010');

-- Office 2013 : http://support.microsoft.com/kb/2786054 
INSERT INTO `officepack_sku` VALUES ('2013','0011','Microsoft Office Professionnel Plus 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0012','Microsoft Office Standard 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0013','Microsoft Office Édition Familial et Business 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0014','Microsoft Office Professionnel 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0015','Microsoft Access 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0016','Microsoft Excel 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0017','Microsoft SharePoint Designer 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0018','Microsoft PowerPoint 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0019','Microsoft Publisher 2013');
INSERT INTO `officepack_sku` VALUES ('2013','001A','Microsoft Outlook 2013');
INSERT INTO `officepack_sku` VALUES ('2013','001B','Microsoft Word 2013');
INSERT INTO `officepack_sku` VALUES ('2013','001C','Microsoft Access Runtime 2013');
INSERT INTO `officepack_sku` VALUES ('2013','001F','Microsoft Office Proofing Tools Kit Compilation 2013');
INSERT INTO `officepack_sku` VALUES ('2013','002F','Microsoft Office famille et étudiant 2013');
INSERT INTO `officepack_sku` VALUES ('2013','003A','Microsoft Project Standard 2013');
INSERT INTO `officepack_sku` VALUES ('2013','003B','Microsoft Project Professionnel 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0044','Microsoft InfoPath 2013');
INSERT INTO `officepack_sku` VALUES ('2013','0057','Microsoft Visio 2013');
INSERT INTO `officepack_sku` VALUES ('2013','00A1','Microsoft OneNote 2013');
INSERT INTO `officepack_sku` VALUES ('2013','00BA','Microsoft Office SharePoint Workspace 2013');
INSERT INTO `officepack_sku` VALUES ('2013','110D','Microsoft Office SharePoint Server 2013');
INSERT INTO `officepack_sku` VALUES ('2013','110F','Microsoft Project Server 2013');
INSERT INTO `officepack_sku` VALUES ('2013','012B','Microsoft Lync 2013');

--
-- Structure de la table `officepack_lang`
--
DROP TABLE IF EXISTS `officepack_lang`;
CREATE TABLE IF NOT EXISTS `officepack_lang` (
  `LCID` varchar(190) DEFAULT NULL,
  `LANG` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`LCID`)
) ENGINE=INNODB;

-- Liste : http://technet.microsoft.com/fr-fr/library/cc179219.aspx
-- Attention le LCID est en héxadécimal
INSERT INTO `officepack_lang` VALUES ('0401','Arabe');
INSERT INTO `officepack_lang` VALUES ('0402','Bulgare');
INSERT INTO `officepack_lang` VALUES ('0804','Chinois (simplifié)');
INSERT INTO `officepack_lang` VALUES ('0404','Chinois');
INSERT INTO `officepack_lang` VALUES ('041A','Croate');
INSERT INTO `officepack_lang` VALUES ('0405','Tchèque');
INSERT INTO `officepack_lang` VALUES ('0406','Danois');
INSERT INTO `officepack_lang` VALUES ('0413','Néerlandais');
INSERT INTO `officepack_lang` VALUES ('0409','Anglais');
INSERT INTO `officepack_lang` VALUES ('0425','Estonien');
INSERT INTO `officepack_lang` VALUES ('040B','Finnois');
INSERT INTO `officepack_lang` VALUES ('040C','Français');
INSERT INTO `officepack_lang` VALUES ('0407','Allemand');
INSERT INTO `officepack_lang` VALUES ('0408','Grec');
INSERT INTO `officepack_lang` VALUES ('040D','Hébreu');
INSERT INTO `officepack_lang` VALUES ('0439','Hindi');
INSERT INTO `officepack_lang` VALUES ('040E','Hongrois');
INSERT INTO `officepack_lang` VALUES ('0410','Italien');
INSERT INTO `officepack_lang` VALUES ('0411','Japonais');
INSERT INTO `officepack_lang` VALUES ('043F','Kazakh');
INSERT INTO `officepack_lang` VALUES ('0412','Coréen');
INSERT INTO `officepack_lang` VALUES ('0426','Letton');
INSERT INTO `officepack_lang` VALUES ('0427','Lituanien');
INSERT INTO `officepack_lang` VALUES ('0414','Norvégien (Bokmal)');
INSERT INTO `officepack_lang` VALUES ('0415','Polonais');
INSERT INTO `officepack_lang` VALUES ('0416','Portugais');
INSERT INTO `officepack_lang` VALUES ('0816','Portugais');
INSERT INTO `officepack_lang` VALUES ('0418','Roumain');
INSERT INTO `officepack_lang` VALUES ('0419','Russe');
INSERT INTO `officepack_lang` VALUES ('081A','Serbe (latin)');
INSERT INTO `officepack_lang` VALUES ('041B','Slovaque');
INSERT INTO `officepack_lang` VALUES ('0424','Slovène');
INSERT INTO `officepack_lang` VALUES ('0C0A','Espagnol');
INSERT INTO `officepack_lang` VALUES ('041D','Suédois');
INSERT INTO `officepack_lang` VALUES ('041E','Thaï');
INSERT INTO `officepack_lang` VALUES ('041F','Turc');
INSERT INTO `officepack_lang` VALUES ('0422','Ukrainien');


--
-- Structure de la table `officepack_type`
--
DROP TABLE IF EXISTS `officepack_type`;
CREATE TABLE IF NOT EXISTS `officepack_type` (
  `REF_ID` varchar(127) DEFAULT NULL,
  `TYPE_VERSION` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`REF_ID`)
) ENGINE=INNODB;

INSERT INTO `officepack_type` VALUES ('0','Licence en volume');
INSERT INTO `officepack_type` VALUES ('1','Vente au détail / OEM');
INSERT INTO `officepack_type` VALUES ('2','Évaluation');
INSERT INTO `officepack_type` VALUES ('5','Téléchargement');


--
-- Structure de la table `officepack_version`
--
DROP TABLE IF EXISTS `officepack_version`;
CREATE TABLE IF NOT EXISTS `officepack_version` (
  `REF_ID` varchar(127) DEFAULT NULL,
  `VERSION` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`REF_ID`)
) ENGINE=INNODB;

INSERT INTO `officepack_version` VALUES ('0','Version avant Beta 1');
INSERT INTO `officepack_version` VALUES ('1','Beta 1');
INSERT INTO `officepack_version` VALUES ('2','Beta 2');
INSERT INTO `officepack_version` VALUES ('3','Version finale candidate 0 (RC0)');
INSERT INTO `officepack_version` VALUES ('4','Version finale candidate 1 (RC1) / OEM Preview Release');
INSERT INTO `officepack_version` VALUES ('9','Version finale (RTM)');
INSERT INTO `officepack_version` VALUES ('A','Service Pack 1');
INSERT INTO `officepack_version` VALUES ('B','Service Pack 2');
INSERT INTO `officepack_version` VALUES ('C','Service Pack 3');
