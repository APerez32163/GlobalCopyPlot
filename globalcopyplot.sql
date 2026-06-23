-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 22-06-2026 a las 07:22:10
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

--
-- Volcado de datos para la tabla `archivo_pedido`
--

INSERT INTO `archivo_pedido` (`ID`, `PEDIDO_ID`, `NOMBRE_ARCHIVO`, `RUTA`) VALUES
(466, 268, 'GlobalCopyplot_Presentacion_IngenieriaSoftware.pptx', 'static/uploads\\impresion\\GlobalCopyplot_Presentacion_IngenieriaSoftware.pptx'),
(467, 268, '5d50bc9f-2752-4d61-9ec3-650a7be69527-1_all_1959.jpg', 'static/uploads\\comprobantes\\5d50bc9f-2752-4d61-9ec3-650a7be69527-1_all_1959.jpg'),
(469, 270, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(470, 271, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(471, 272, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(472, 273, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(473, 274, 'GlobalCopyplotPresentacionProyecto.pptx', 'static/uploads\\impresion\\GlobalCopyplotPresentacionProyecto.pptx'),
(474, 274, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(475, 275, 'GlobalCopyplotPresentacionProyecto.pptx', 'static/uploads\\impresion\\GlobalCopyplotPresentacionProyecto.pptx'),
(476, 276, 'PNFI-Trayecto2Seccion2.pdf', 'static/uploads\\impresion\\PNFI-Trayecto2Seccion2.pdf'),
(477, 277, 'GlobalCopyplot_Presentacion_IngenieriaSoftware.pptx', 'static/uploads\\impresion\\GlobalCopyplot_Presentacion_IngenieriaSoftware.pptx'),
(478, 277, 'GlobalCopyplotPresentacionProyecto.pptx', 'static/uploads\\impresion\\GlobalCopyplotPresentacionProyecto.pptx'),
(479, 277, 'Captura1.PNG', 'static/uploads\\comprobantes\\Captura1.PNG'),
(480, 278, 'Captura1.PNG', 'static/uploads\\impresion\\Captura1.PNG'),
(481, 278, '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG', 'static/uploads\\comprobantes\\0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG'),
(482, 279, '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG', 'static/uploads\\impresion\\0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG'),
(483, 279, 'Captura1.PNG', 'static/uploads\\comprobantes\\Captura1.PNG'),
(484, 280, 'Captura1.PNG', 'static/uploads\\impresion\\Captura1.PNG'),
(485, 280, '0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG', 'static/uploads\\comprobantes\\0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000.PNG');

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
  `PRECIO_UNITARIO` decimal(10,2) DEFAULT NULL,
  `IVA` decimal(3,2) DEFAULT NULL,
  `TOTAL_PAGO` decimal(10,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `detalle`
--

INSERT INTO `detalle` (`ID`, `CANTIDAD`, `SUBTOTAL`, `PEDIDO_ID`, `PRECIO_UNITARIO`, `IVA`, `TOTAL_PAGO`) VALUES
(131, 11, 550.00, 268, 50.00, 0.00, 550.00),
(132, 65, 3250.00, 270, 50.00, 0.00, 3250.00),
(133, 65, 3250.00, 271, 50.00, 0.00, 3250.00),
(134, 65, 3250.00, 272, 50.00, 0.00, 3250.00),
(135, 65, 3250.00, 273, 50.00, 0.00, 3250.00),
(136, 77, 6800.00, 274, 50.00, 0.00, 6800.00),
(137, 12, 700.00, 275, 50.00, 0.00, 700.00),
(138, 65, 6000.00, 276, 50.00, 0.00, 6000.00),
(139, 23, 1700.00, 277, 50.00, 0.00, 1700.00),
(140, 1, 50.00, 278, 50.00, 0.00, 50.00),
(141, 1, 200.00, 279, 200.00, 0.00, 200.00),
(142, 1, 100.00, 280, 100.00, 0.00, 100.00);

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

--
-- Volcado de datos para la tabla `pedido`
--

INSERT INTO `pedido` (`ID`, `FECHA`, `ESTADO`, `TOTAL`, `TAMANO`, `PAGINAS_COLOR`, `ID_USUARIO`, `FECHA_RETIRO`, `HORA_RETIRO`, `CODIGO_TICKET`, `COMENTARIOS`, `VISTO_ADMIN`, `VISTO_OPERADOR`, `PAGINAS`, `SERVICIO_ID`, `REFERENCIA_PAGO`, `DETALLE_ARCHIVOS`) VALUES
(268, '2026-06-21 22:48:13', 'Listo', 550.00, 'Carta', NULL, 19, '2026-06-23', '09:00:00', 'TICK-268-7M0Z', '', 1, 1, 11, 3, '555556656666666', NULL),
(270, '2026-06-22 00:16:11', 'Cancelado', 3250.00, 'Carta', NULL, 19, '2026-06-24', '10:00:00', NULL, '', 0, 0, 65, 14, NULL, NULL),
(271, '2026-06-22 00:17:24', 'Cancelado', 3250.00, 'Carta', NULL, 19, '2026-06-23', '08:30:00', NULL, '', 0, 0, 65, 14, NULL, NULL),
(272, '2026-06-22 00:20:54', 'Cancelado', 3250.00, 'Carta', NULL, 19, '2026-06-23', '10:00:00', NULL, 'tengo hijos', 0, 0, 65, 14, NULL, NULL),
(273, '2026-06-22 00:23:13', 'Cancelado', 3250.00, 'Carta', NULL, 19, '2026-06-23', '09:30:00', NULL, 'tengo hijos', 0, 0, 65, 14, NULL, NULL),
(274, '2026-06-22 00:27:33', 'Cancelado', 6800.00, NULL, NULL, 19, '2026-06-23', '10:30:00', NULL, '', 0, 0, 77, NULL, NULL, '[{\"nombre\": \"GlobalCopyplotPresentacionProyecto.pptx\", \"paginas\": 12, \"servicio_id\": 14, \"tamano\": \"Carta\", \"paginas_color\": \"4\"}, {\"nombre\": \"PNFI-Trayecto2Seccion2.pdf\", \"paginas\": 65, \"servicio_id\": 14, \"tamano\": \"Carta\", \"paginas_color\": \"55\"}]'),
(275, '2026-06-22 00:32:16', 'Cancelado', 700.00, 'Carta', '2', 19, '2026-06-23', '10:00:00', NULL, '', 0, 0, 12, 14, NULL, NULL),
(276, '2026-06-22 00:32:54', 'Cancelado', 6000.00, 'Carta', '55', 19, '2026-06-23', '10:30:00', NULL, '', 0, 0, 65, 14, NULL, NULL),
(277, '2026-06-22 00:43:43', 'Pago confirmado', 1700.00, NULL, NULL, 19, '2026-06-23', '09:30:00', 'TICK-277-AZC9', '', 1, 0, 23, NULL, '000000000000000', '[{\"nombre\": \"GlobalCopyplot_Presentacion_IngenieriaSoftware.pptx\", \"paginas\": 11, \"servicio_id\": 14, \"tamano\": \"Carta\", \"paginas_color\": \"5\"}, {\"nombre\": \"GlobalCopyplotPresentacionProyecto.pptx\", \"paginas\": 12, \"servicio_id\": 14, \"tamano\": \"Carta\", \"paginas_color\": \"6\"}]'),
(278, '2026-06-22 00:56:06', 'Pago confirmado', 50.00, 'Carta', NULL, 19, '2026-06-23', '17:30:00', 'TICK-278-8M9P', '', 1, 0, 1, 3, '444444444444444', NULL),
(279, '2026-06-22 00:57:14', 'Entregado', 200.00, 'Doble carta', NULL, 19, '2026-06-23', '07:00:00', 'TICK-279-J73U', '', 1, 0, 1, 4, '455555555555555', NULL),
(280, '2026-06-22 00:58:05', 'Pago confirmado', 100.00, 'Extra oficio', NULL, 19, '2026-06-29', '09:30:00', 'TICK-280-GC1N', '', 1, 0, 1, 3, '125222222222222', NULL);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `servicio_impresion`
--

CREATE TABLE `servicio_impresion` (
  `ID` int(11) NOT NULL,
  `TITULO` varchar(100) NOT NULL,
  `DESCRIPCION` text DEFAULT NULL,
  `ACTIVO` tinyint(1) DEFAULT 1,
  `ES_MIXTO` tinyint(1) DEFAULT 0,
  `SERVICIO_BN_ID` int(11) DEFAULT NULL,
  `SERVICIO_COLOR_ID` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `servicio_impresion`
--

INSERT INTO `servicio_impresion` (`ID`, `TITULO`, `DESCRIPCION`, `ACTIVO`, `ES_MIXTO`, `SERVICIO_BN_ID`, `SERVICIO_COLOR_ID`) VALUES
(3, 'Impresión B/N Bond', 'Impresión blanco y negro.', 1, 0, NULL, NULL),
(4, 'Impresión Laser Color Bond', 'Impresión a color.', 1, 0, NULL, NULL),
(14, 'Impresión Mixta Bond', 'Impresión a color y blanco y negro de forma mixta.', 1, 1, NULL, NULL);

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
(9, 3, 'Carta', 50.00, 0.00),
(10, 3, 'Oficio', 80.00, 0.00),
(11, 3, 'Extra oficio', 100.00, 0.00),
(12, 3, 'Doble carta', 160.00, 0.00),
(77, 4, 'Carta', 100.00, 100.00),
(78, 4, 'Oficio', 130.00, 130.00),
(79, 4, 'Extra oficio', 150.00, 150.00),
(80, 4, 'Doble carta', 200.00, 200.00),
(81, 14, 'Carta', 50.00, 100.00),
(82, 14, 'Oficio', 80.00, 130.00),
(83, 14, 'Extra oficio', 100.00, 150.00),
(84, 14, 'Doble carta', 160.00, 200.00);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `usuario`
--

CREATE TABLE `usuario` (
  `ID` int(11) NOT NULL,
  `ID_USUARIO` varchar(10) DEFAULT NULL,
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
  ADD PRIMARY KEY (`ID`),
  ADD KEY `SERVICIO_BN_ID` (`SERVICIO_BN_ID`),
  ADD KEY `SERVICIO_COLOR_ID` (`SERVICIO_COLOR_ID`);

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
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=486;

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
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=143;

--
-- AUTO_INCREMENT de la tabla `pedido`
--
ALTER TABLE `pedido`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=281;

--
-- AUTO_INCREMENT de la tabla `servicio_impresion`
--
ALTER TABLE `servicio_impresion`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=16;

--
-- AUTO_INCREMENT de la tabla `servicio_impresion_tamano`
--
ALTER TABLE `servicio_impresion_tamano`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=85;

--
-- AUTO_INCREMENT de la tabla `usuario`
--
ALTER TABLE `usuario`
  MODIFY `ID` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=26;

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
-- Filtros para la tabla `servicio_impresion`
--
ALTER TABLE `servicio_impresion`
  ADD CONSTRAINT `servicio_impresion_ibfk_1` FOREIGN KEY (`SERVICIO_BN_ID`) REFERENCES `servicio_impresion` (`ID`),
  ADD CONSTRAINT `servicio_impresion_ibfk_2` FOREIGN KEY (`SERVICIO_COLOR_ID`) REFERENCES `servicio_impresion` (`ID`);

--
-- Filtros para la tabla `servicio_impresion_tamano`
--
ALTER TABLE `servicio_impresion_tamano`
  ADD CONSTRAINT `servicio_impresion_tamano_ibfk_1` FOREIGN KEY (`SERVICIO_ID`) REFERENCES `servicio_impresion` (`ID`) ON DELETE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
