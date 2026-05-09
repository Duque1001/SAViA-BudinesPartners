// Estilos del header
import "../styles/Header.css";

// Logo por defecto
import defaultLogo from "../assets/images/logo-savios.png";
//import defaultLogo from "../assets/images/logo-comfama.png"; //CAMBIO COMFAMA

// Funciones para obtener datos de sesión
import { getName, getUniqueName } from "../config/session";

// Hook de traducciones
import { useTranslation } from "react-i18next";

// Hooks de React
import { useEffect, useState } from "react";

// Props del componente Header
interface HeaderProps {
  username: string;          // Usuario base
  displayName?: string;      // Nombre visible opcional
  onLogout: () => void;      // Función cerrar sesión
  onToggleMenu?: () => void; // Abrir/cerrar menú lateral
}

// Componente principal
export default function Header({
  username,
  displayName,
  onLogout,
  onToggleMenu,
}: HeaderProps) {
  // Hook de traducción (namespaces: common y chat)
  const { t } = useTranslation(["common", "chat"]);

  // Obtener datos de sesión almacenados
  const storedName = getName();
  const storedUnique = getUniqueName();

  // Prioridad del nombre a mostrar
  const nameToShow = displayName || storedName || storedUnique || username;

  // Estado del logo (localStorage o por defecto)
  const [logo, setLogo] = useState<string>(
    localStorage.getItem("app_logo") || defaultLogo
  );

  // Escucha cambios del logo en tiempo real
  useEffect(() => {
    const handleLogoUpdate = () => {
      setLogo(localStorage.getItem("app_logo") || defaultLogo);
    };

    // Evento personalizado para actualizar logo
    window.addEventListener("logoUpdated", handleLogoUpdate);

    // Limpieza al desmontar
    return () => {
      window.removeEventListener("logoUpdated", handleLogoUpdate);
    };
  }, []);

  return (
    <header className="header">
      {/* Lado izquierdo: menú + logo */}
      <div className="header__left">
        <button
          className="menu-toggle"
          aria-label={t("openMenu", { ns: "common" })} // accesibilidad
          onClick={onToggleMenu} // abre/cierra sidenav
        >
          ☰
        </button>

        {/* Logo dinámico 
        <img src={logo} alt="SAVIOS" className="logo" />*/}
        <div className="logo-container">
          <img src={logo} alt="SAVIOS" className="header-logo" />
        </div>
      </div>

      {/* Lado derecho: usuario + logout */}
      <div className="header__right">
        <span className="welcome">
          {t("hello", { ns: "common" })}: <strong>{nameToShow}</strong>
        </span>

        {/* Botón cerrar sesión */}
        <button className="btn btn--logout" onClick={onLogout}>
          {t("logout", { ns: "chat" })}
        </button>
      </div>
    </header>
  );
}