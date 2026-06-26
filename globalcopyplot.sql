-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 26-06-2026 a las 21:48:14
-- Versión del servidor: 10.4.32-MariaDB
-- Versión de PHP: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `globalcopyplot`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `archivo_pedido`
--

CREATE TABLE `archivo_pedido` (
  `ID` int(11) NOT NULL,
  `PEDIDO_ID` int(11) NOT NULL,
  `NOMBRE_ARCHIVO` varchar(255) NOT NULL,
  `RUTA` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `catalogo`
--

CREATE TABLE `catalogo` (
  `ID` int(11) NOT NULL,
  `IMAGEN` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `catalogo`
--

INSERT INTO `catalogo` (`ID`, `IMAGEN`) VALUES
(1, '1.png'),
(10, '10.png'),
(11, '11.png'),
(12, '12.png'),
(13, '13.png'),
(16, '14.png'),
(2, '2.png'),
(3, '3.png'),
(4, '4.png'),
(5, '5.png'),
(6, '6.png'),
(7, '7.png'),
(8, '8.png'),
(9, '9.png');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `configuracion`
--

CREATE TABLE `configuracion` (
  `ID` int(11) NOT NULL,
  `CLAVE` varchar(100) NOT NULL,
  `VALOR` text DEFAULT NULL,
  `DESCRIPCION` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `configuracion`
--

INSERT INTO `configuracion` (`ID`, `CLAVE`, `VALOR`, `DESCRIPCION`) VALUES
(1, 'nombre_negocio', 'GlobalCopyPlot', 'Nombre del negocio'),
(2, 'direccion', 'Los Teques - Calle Ribas', 'Dirección física'),
(3, 'telefono_negocio', '04122826743', 'Teléfono de contacto'),
(4, 'whatsapp', '584122826743', 'Número WhatsApp (formato internacional sin +)'),
(5, 'instagram', 'globalcopyplot', 'Usuario de Instagram'),
(6, 'email_negocio', 'GlobalCopyPlot@gmail.com', 'Correo electrónico del negocio'),
(7, 'pago_movil_banco', '0102 - Banco de Venezuela', 'Banco para PagoMóvil'),
(8, 'pago_movil_cedula', '32163794', 'Cédula para PagoMóvil'),
(9, 'pago_movil_telefono', '0412-2826746', 'Teléfono para PagoMóvil'),
(13, 'correo_notificaciones', 'SupportGlobalCopyPlot@gmail.com', 'Correo desde el que se envían notificaciones'),
(14, 'notificar_pedido_listo', '1', 'Notificar cuando un pedido esté listo (1=Sí, 0=No)'),
(15, 'politicas', 'Hola', 'Texto de políticas y términos'),
(16, 'respaldos_automaticos', '1', 'Activar respaldos automáticos (1=Sí, 0=No)'),
(17, 'frecuencia_respaldo', 'semanal', 'Frecuencia de respaldo: diario, semanal, mensual');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `detalle`
--

CREATE TABLE `detalle` (
  `ID` int(11) NOT NULL,
  `CANTIDAD` int(5) DEFAULT NULL,
  `SUBTOTAL` decimal(10,2) DEFAULT NULL,
  `PEDIDO_ID` int(11) NOT NULL,
  `PRECIO_UNITARIO` decimal(10,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pedido`
--

CREATE TABLE `pedido` (
  `ID` int(11) NOT NULL,
  `FECHA` datetime DEFAULT current_timestamp(),
  `ESTADO` varchar(20) DEFAULT 'Pendiente',
  `TOTAL` decimal(10,2) DEFAULT 0.00,
  `TAMANO` varchar(20) DEFAULT NULL,
  `PAGINAS_COLOR` varchar(255) DEFAULT NULL,
  `ID_USUARIO` int(11) DEFAULT NULL,
  `FECHA_RETIRO` date DEFAULT NULL,
  `HORA_RETIRO` time DEFAULT NULL,
  `CODIGO_TICKET` varchar(20) DEFAULT NULL,
  `COMENTARIOS` text DEFAULT NULL,
  `VISTO_ADMIN` tinyint(1) NOT NULL DEFAULT 0,
  `VISTO_OPERADOR` tinyint(1) NOT NULL DEFAULT 0,
  `PAGINAS` int(11) DEFAULT NULL,
  `SERVICIO_ID` int(11) DEFAULT NULL,
  `REFERENCIA_PAGO` varchar(100) DEFAULT NULL,
  `DETALLE_ARCHIVOS` longtext CHARACTER SET utf8mb4 COLLATE utf8mb4_bin DEFAULT NULL CHECK (json_valid(`DETALLE_ARCHIVOS`))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `servicio_impresion`
--

CREATE TABLE `servicio_impresion` (
  `ID` int(11) NOT NULL,
  `TITULO` varchar(100) NOT NULL,
  `DESCRIPCION` text DEFAULT NULL,
  `ACTIVO` tinyint(1) DEFAULT 1,
  `ES_MIXTO` tinyint(1) DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `servicio_impresion`
--

INSERT INTO `servicio_impresion` (`ID`, `TITULO`, `DESCRIPCION`, `ACTIVO`, `ES_MIXTO`) VALUES
(21, 'Impresión Mixta Bond', 'Impresión a color y blanco y negro para documentos que requieran paginas especificas a color', 1, 1),
(22, 'Impresión B/N Bond', 'Impresión para documentos unicamente blanco y negro', 1, 0),
(23, 'Impresion Laser Color Bond', 'Impresión a color (todo el documento)', 1, 0);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `servicio_impresion_tamano`
--

CREATE TABLE `servicio_impresion_tamano` (
  `ID` int(11) NOT NULL,
  `SERVICIO_ID` int(11) NOT NULL,
  `NOMBRE` varchar(50) NOT NULL,
  `PRECIO_BN` decimal(10,2) NOT NULL,
  `PRECIO_COLOR` decimal(10,2) NOT NULL DEFAULT 0.00
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `servicio_impresion_tamano`
--

INSERT INTO `servicio_impresion_tamano` (`ID`, `SERVICIO_ID`, `NOMBRE`, `PRECIO_BN`, `PRECIO_COLOR`) VALUES
(98, 21, 'Carta', 80.00, 100.00),
(99, 21, 'Oficio', 120.00, 140.00),
(100, 21, 'Extra oficio', 160.00, 180.00),
(101, 21, 'Doble Carta', 200.00, 220.00),
(102, 22, 'Carta', 80.00, 80.00),
(103, 22, 'Oficio', 120.00, 120.00),
(104, 22, 'Extra oficio', 160.00, 160.00),
(105, 22, 'Doble carta', 200.00, 200.00),
(106, 23, 'Carta', 100.00, 100.00),
(107, 23, 'Oficio', 140.00, 140.00),
(108, 23, 'Extra oficio', 180.00, 180.00),
(109, 23, 'Doble carta', 220.00, 220.00);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `usuario`
--

CREATE TABLE `usuario` (
  `ID` int(11) NOT NULL,
  `ID_USUARIO` varchar(15) NOT NULL,
  `NOMBRE` varchar(50) NOT NULL,
  `APELLIDO` varchar(50) NOT NULL,
  `EMAIL` varchar(120) NOT NULL,
  `CONTRASEÑA` varchar(255) NOT NULL,
  `TELEFONO` varchar(11) NOT NULL,
  `CONFIRMADO` tinyint(1) NOT NULL DEFAULT 0,
  `ES_ADMIN` tinyint(1) NOT NULL DEFAULT 0,
  `ES_OPERADOR` tinyint(1) NOT NULL DEFAULT 0,
  `PREGUNTA1` varchar(200) DEFAULT NULL,
  `RESPUESTA1` varchar(200) DEFAULT NULL,
  `PREGUNTA2` varchar(200) DEFAULT NULL,
  `RESPUESTA2` varchar(200) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `usuario`
--

INSERT INTO `usuario` (`ID`, `ID_USUARIO`, `NOMBRE`, `APELLIDO`, `EMAIL`, `CONTRASEÑA`, `TELEFONO`, `CONFIRMADO`, `ES_ADMIN`, `ES_OPERADOR`, `PREGUNTA1`, `RESPUESTA1`, `PREGUNTA2`, `RESPUESTA2`) VALUES
(18, '00000000', 'Prueba', 'Operador', 'andresperez32163@gmail.com', 'scrypt:32768:8:1$Av0UTck9Oy74TFF5$889113c6c2d26297e109ea79852e0e18d150689564653910b8d801eeed6749fe087f3984da000ac8e1d0988cc2ba1d315a960afec1b68f055827671b4d217ada', '00000000000', 1, 0, 1, 'color', 'ROJO', 'comida', 'PASTA'),
(19, '11111111', 'Prueba', 'Admin', 'andreseduardo32163@gmail.com', 'scrypt:32768:8:1$T0B5k0eET9Q4QtWS$9f4c563d170e1be3c2dd95c122b31a1d28deea41895ace38febe77ec7d7dbc1536a1b1a70d36af7c3d09320b388f5b93a2236b43eeb1f52a2080379afd7fcae6', '11111111111', 1, 1, 0, 'color', 'ROJO', 'comida', 'PASTA'),
(20, '22222222', 'Prueba', 'Usuario Cliente', 'kenyersoncrespo6@gmail.com', 'scrypt:32768:8:1$i1jzZAWBIIDfJA5J$f2ca81b00b2758323a510d52d61871107d07e82653de499dd6fe535fca85f24ac98ff3f5f6d042ebbd385dbb8bd4a98b519ffbb3737756e60f1b823140d66d71', '22222222222', 1, 0, 0, 'color', 'ROJO', 'comida', 'PASTA');

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `archivo_pedido`
--
ALTER TABLE `archivo_pedido`
  ADD PRIMARY KEY (`ID`),
  ADD KEY `PEDIDO_ID` (`PEDIDO_ID`);

--
-- Indices de la tabla `catalogo`
--
ALTER TABLE `catalogo`
  ADD PRIMARY KEY (`ID`),
  ADD UNIQUE KEY `IMAGEN` (`IMAGEN`);

--
-- Indices de la tabla `configuracion`
--
ALTER TABLE `configuracion`
  ADD PRIMARY KEY (`ID`),
  ADD UNIQUE KEY `CLAVE` (`CLAVE`);

--
-- Indices de la tabla `detalle`
--
ALTER TABLE `detalle`
  ADD PRIMARY KEY (`ID`),
  ADD KEY `ID_PEDIDO` (`PEDIDO_ID`);

--
-- Indices de la tabla `pedido`
--
ALTER TABLE `pedido`
  ADD PRIMARY KEY (`ID`),
  ADD UNIQUE KEY `CODIGO_TICKET` (`CODIGO_TICKET`),
  ADD KEY `ID_USUARIO` (`ID_USUARIO`),
  ADD KEY `SERVICIO_ID` (`SERVICIO_ID`);

--
-- Indices de la tabla `servicio_impresion`
--
ALTER TABLE `servicio_impresion`
  ADD PRIMARY KEY (`ID`);

--
-- Indices de la tabla `servicio_impresion_tamano`
--
ALTER TABLE `servicio_impresion_tamano`
  ADD PRIMARY KEY (`ID`),
  ADD KEY `SERVICIO_ID` (`SERVICIO_ID`);

--
-- Indices de la tabla `usuario`
--
ALTER TABLE `usuario`
  ADD PRIMARY KEY (`ID`),
  ADD UNIQUE KEY `EMAIL` (`EMAIL`),
  ADD UNIQUE KEY `TELEFONO` (`TELEFONO`),
  ADD UNIQUE KEY `ID_USUARIO` (`ID_USUARIO`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `archivo_pedido`
--
ALTER TABLE `archivo_pedido`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=774;

--
-- AUTO_INCREMENT de la tabla `catalogo`
--
ALTER TABLE `catalogo`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=17;

--
-- AUTO_INCREMENT de la tabla `configuracion`
--
ALTER TABLE `configuracion`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=27;

--
-- AUTO_INCREMENT de la tabla `detalle`
--
ALTER TABLE `detalle`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=172;

--
-- AUTO_INCREMENT de la tabla `pedido`
--
ALTER TABLE `pedido`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=396;

--
-- AUTO_INCREMENT de la tabla `servicio_impresion`
--
ALTER TABLE `servicio_impresion`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=24;

--
-- AUTO_INCREMENT de la tabla `servicio_impresion_tamano`
--
ALTER TABLE `servicio_impresion_tamano`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=110;

--
-- AUTO_INCREMENT de la tabla `usuario`
--
ALTER TABLE `usuario`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=30;

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `archivo_pedido`
--
ALTER TABLE `archivo_pedido`
  ADD CONSTRAINT `archivo_pedido_ibfk_1` FOREIGN KEY (`PEDIDO_ID`) REFERENCES `pedido` (`ID`);

--
-- Filtros para la tabla `pedido`
--
ALTER TABLE `pedido`
  ADD CONSTRAINT `pedido_ibfk_1` FOREIGN KEY (`ID_USUARIO`) REFERENCES `usuario` (`ID`),
  ADD CONSTRAINT `pedido_ibfk_2` FOREIGN KEY (`SERVICIO_ID`) REFERENCES `servicio_impresion` (`ID`);

--
-- Filtros para la tabla `servicio_impresion_tamano`
--
ALTER TABLE `servicio_impresion_tamano`
  ADD CONSTRAINT `servicio_impresion_tamano_ibfk_1` FOREIGN KEY (`SERVICIO_ID`) REFERENCES `servicio_impresion` (`ID`) ON DELETE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
