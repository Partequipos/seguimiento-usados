/**
 * Componente de botón de inicio de sesión
 */

import React from "react";
import { useAuth } from "../context/AuthContext";
import { LogIn, LogOut, Loader2 } from "lucide-react";

const LoginButton: React.FC = () => {
  const { isAuthenticated, account, isLoading, login, logout } = useAuth();

  if (isLoading) {
    return (
      <button
        disabled
        className="flex items-center gap-2 px-4 py-2 bg-brand-gray-soft text-brand-gray-light rounded-lg cursor-not-allowed border border-brand-gray-border"
      >
        <Loader2 className="w-5 h-5 animate-spin" />
        Cargando...
      </button>
    );
  }

  if (isAuthenticated && account) {
    return (
      <div className="flex items-center gap-3">
        <div className="text-right hidden sm:block">
          <p className="text-sm font-semibold text-brand-ink">{account.name}</p>
          <p className="text-xs text-brand-gray-light">{account.username}</p>
        </div>
        <button
          onClick={logout}
          className="flex items-center gap-2 px-4 py-2 bg-brand-gray hover:bg-brand-ink text-brand-white rounded-lg transition-colors font-medium"
        >
          <LogOut className="w-5 h-5" />
          Cerrar Sesión
        </button>
      </div>
    );
  }

  return (
    <button
      onClick={login}
      className="btn-brand"
    >
      <LogIn className="w-5 h-5" />
      Iniciar Sesión con Microsoft
    </button>
  );
};

export default LoginButton;
