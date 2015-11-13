/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.27-community-nt : Database - usuarios
*********************************************************************
Server version : 5.0.27-community-nt
*/

SET NAMES utf8;

SET SQL_MODE='';

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `usuariosiconosppal` */

CREATE TABLE `usuariosiconosppal` (
  `codusu` smallint(6) NOT NULL,
  `aplicacion` varchar(30) NOT NULL,
  `PuntoMenu` int(11) NOT NULL,
  `icono` smallint(6) NOT NULL,
  `TextoOrigen` varchar(100) NOT NULL,
  `TextoVisible` varchar(100) NOT NULL,
  PRIMARY KEY  (`codusu`,`aplicacion`,`PuntoMenu`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `usuariosiconosppal` */

insert into `usuariosiconosppal` (`codusu`,`aplicacion`,`PuntoMenu`,`icono`,`TextoOrigen`,`TextoVisible`) values (9,'ariconta',12,12,'Conceptos','Conceptos'),(9,'ariconta',13,13,'Tipos de I.V.A.','Tipos de I.V.A.'),(9,'ariconta',14,14,'Bancos','Bancos'),(9,'ariconta',16,16,'Formas de pago','Formas de pago'),(9,'ariconta',18,18,'Plan contable','Plan contable'),(9,'ariconta',19,19,'Asientos','Asientos'),(9,'ariconta',21,21,'Diario','Diario'),(9,'ariconta',28,28,'Facturas emitidas','Facturas emitidas'),(12,'ariconta',14,14,'Bancos','Bancos'),(12,'ariconta',19,19,'Asientos','Asientos'),(12,'ariconta',28,28,'Facturas emitidas','Facturas emitidas'),(13,'ariconta',13,13,'Tipos de I.V.A.','Tipos de I.V.A.'),(1005,'ariconta',11,11,'Asientos predefinidos','Asientos predefinidos'),(1005,'ariconta',12,12,'Conceptos','Conceptos'),(1005,'ariconta',14,14,'Bancos','Bancos'),(1005,'ariconta',17,17,'Agentes','Agentes'),(1005,'ariconta',21,21,'Diario','Diario');

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
