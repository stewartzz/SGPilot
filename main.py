#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SGPilot v5.2
- Rebranded: SGPilot
- Duas abas: SGP + Papervines com temas visuais distintos
- SGP: dark navy + verde/amarelo (combina com logo SGPilot)
- Papervines: claro + roxo (combina com estética Papervines)
- Transição animada suave entre abas (10 steps fade)
- Hovers coerentes por tema
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import sys
import json
import time
import threading
import re
import logging
import urllib.request
from datetime import datetime
from pathlib import Path

# ── Logging ─────────────────────────────────────────────────
_log_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
LOG_FILE = _log_dir / "sgp_auto.log"
try:
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
except Exception:
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
    )
log = logging.getLogger("SGPAuto")

# ── Imports opcionais ────────────────────────────────────────

try:
    import keyboard
    KEYBOARD_OK = True
except ImportError:
    KEYBOARD_OK = False

try:
    import pyautogui
    import pyperclip
    pyautogui.FAILSAFE = True
    PYAUTOGUI_OK = True
except ImportError:
    PYAUTOGUI_OK = False

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

# ════════════════════════════════════════════════════════════
#  CONSTANTES
# ════════════════════════════════════════════════════════════

APP_VERSION = "5.2.0"

# PyInstaller: __file__ aponta para temp, mas os logos ficam junto ao .exe
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    APP_DIR = Path(__file__).parent
    BUNDLE_DIR = APP_DIR

CONFIG_FILE = APP_DIR / "config.json"
LOGO_ICON   = APP_DIR / "logo.png"        # Ícone da taskbar
LOGO_HEADER = APP_DIR / "logo2.png"       # SGPilot header logo
LOGO3       = APP_DIR / "logo3.png"       # Logo centralizado no header
LOGO_PV     = APP_DIR / "logo_pv.png"     # Logo Papervines

# Fallback: tenta no bundle se não achou na pasta do exe
if not LOGO_ICON.exists():
    LOGO_ICON = BUNDLE_DIR / "logo.png"
if not LOGO_HEADER.exists():
    LOGO_HEADER = BUNDLE_DIR / "logo2.png"
if not LOGO3.exists():
    LOGO3 = BUNDLE_DIR / "logo3.png"
if not LOGO_PV.exists():
    LOGO_PV = BUNDLE_DIR / "logo_pv.png"

# ── Temas visuais ──────────────────────────────────────────

# SGP: cinza idêntico ao site SGP (TSMX)
THEME_SGP = {
    "bg":           "#212121",
    "bg_card":      "#2a2a2a",
    "bg_card_hover":"#333333",
    "bg_header":    "#1a1a1a",
    "bg_input":     "#2a2a2a",
    "accent":       "#28a745",       # verde Cadastrar
    "accent_dim":   "#218838",
    "accent_dark":  "#1e7e34",
    "accent_hover": "#2fbd4f",
    "blue":         "#337AB7",       # azul Cadastrar Ocorrência
    "blue_hover":   "#3D8FD4",
    "red":          "#CC3333",
    "red_dim":      "#3D1515",
    "yellow":       "#E8A735",
    "text":         "#cacacc",
    "text_dim":     "#888888",
    "border":       "#3a3a3a",
    "font":         "Roboto",
    "font_fb":      "Segoe UI",
    "tab_active":   "#424242",
    "tab_inactive": "#2a2a2a",
    "switch_on":    "#28a745",
}

# Papervines: claro, clean, roxo/violeta
THEME_PV = {
    "bg":           "#F7F7FB",
    "bg_card":      "#FFFFFF",
    "bg_card_hover":"#EEEDF5",
    "bg_header":    "#FFFFFF",
    "bg_input":     "#F0EFF5",
    "accent":       "#7C3AED",
    "accent_dim":   "#6D28D9",
    "accent_dark":  "#EDE9FE",
    "accent_hover": "#DDD6FE",
    "blue":         "#7C3AED",
    "blue_hover":   "#6D28D9",
    "red":          "#EF4444",
    "red_dim":      "#FEE2E2",
    "yellow":       "#F59E0B",
    "text":         "#1F2937",
    "text_dim":     "#9CA3AF",
    "border":       "#E5E7EB",
    "font":         "Roboto",
    "font_fb":      "Segoe UI",
    "tab_active":   "#EDE9FE",
    "tab_inactive": "#FFFFFF",
    "switch_on":    "#7C3AED",
}

# Alias ativo (usado pelos editores/janelas que referenciam THEME)
THEME = THEME_SGP

# Tipos de ação disponíveis — uma bind pode ter VÁRIOS ao mesmo tempo
TIPOS_ACAO = [
    ("text",          "Enviar texto no chat"),
    ("text_ocr",      "Texto + Número (lê do SGP via HTML)"),
    ("sgp_ocorrencia","Automação SGP (preenche formulário)"),
    ("link_pagamento","Link de pagamento (2 etapas)"),
]

# IDs reais dos campos HTML do SGP (inspecionados via DevTools)
SGP_IDS = {
    "tipo_container":     "select2-id_tipo-container",
    "origem_container":   "select2-id_metodo-container",
    "contrato_container": "select2-id_clientecontrato-container",
    "conteudo":           "id_conteudo",
    "gerar_os":           "id_os",
    "cadastrar":          "btacao",
}

# SGPs que usam "finan" ao invés de "sus" para tipo financeiro
SGPS_TIPO_FINAN = ["supersonic"]   # VMA = supersonic.sgp.tsmx.com.br

DEFAULT_CONFIG = {
    "version": APP_VERSION,
    "sgp": {
        "debug_port": 9222,
        "delay_ms":   150,
    },
    "papervines": {
        "enabled": True,
        "bind_key": "F9",
        "saudacao": "Olá! Tudo bem? Meu nome é Atendente, como posso te ajudar?",
        "delay_entre_clientes_ms": 1500,
        "transfer_dept": "Entrada Central",
        "transfer_bind_key": "F10",
    },
    "binds": [
        {
            "id": "f1", "key": "F1", "enabled": True, "cor": "#1D9E75",
            "name": "OFF LOSI/LOBI",
            "types": ["text_ocr"],
            "message": "OFF LOSI/LOBI\n\n{numero}",
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        },
        {
            "id": "f2", "key": "F2", "enabled": True, "cor": "#1D9E75",
            "name": "OFF FTTx com potência",
            "types": ["text_ocr"],
            "message": "OFF COM POTÊNCIA FTTx (feito procedimentos porém sem sucesso)\n\n{numero}",
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        },
        {
            "id": "f3", "key": "F3", "enabled": True, "cor": "#1D9E75",
            "name": "Offline Dying-gasp",
            "types": ["text_ocr"],
            "message": "OFFLINE Dying-gasp\n\n{numero}",
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        },
        {
            "id": "f4", "key": "F4", "enabled": True, "cor": "#1D9E75",
            "name": "Sinal atenuado",
            "types": ["text_ocr"],
            "message": "Sinal atenuado -27.00 FTTx\n\n{numero}",
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        },
        {
            "id": "f6", "key": "F6", "enabled": True, "cor": "#185FA5",
            "name": "Comprovante (SGP completo)",
            "types": ["sgp_ocorrencia"],
            "message": "Comprovante de pagamento referente ao mês {mes}",
            "sgp_tipo_filtro": "sus", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        },
        {
            "id": "f7", "key": "F7", "enabled": True, "cor": "#534AB7",
            "name": "Link de pagamento",
            "types": ["link_pagamento"],
            "message": (
                "Nesse link possui todos os dados para pagamento da fatura "
                "Pix, Código de Barras e QR Code, e também um botão para acesso à fatura.\n"
                "*Pedimos que verifique os dados antes de realizar o pagamento 😉*\n"
                "\n*Link para pagamento abaixo:*\n{link}"
            ),
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False, "sgp_auto_cadastrar_os": False
        }
    ]
}

# ════════════════════════════════════════════════════════════
#  CONFIG MANAGER
# ════════════════════════════════════════════════════════════

class ConfigManager:
    def __init__(self):
        self.data = self._load()

    def _load(self):
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    d = json.load(f)
                # Migra binds antigas que usam "type" (string) para "types" (lista)
                for b in d.get("binds", []):
                    if "type" in b and "types" not in b:
                        b["types"] = [b.pop("type")]
                    elif "types" not in b:
                        b["types"] = ["text"]
                    # Migra binds sem sgp_auto_cadastrar (padrão: False)
                    if "sgp_auto_cadastrar" not in b:
                        b["sgp_auto_cadastrar"] = False
                    # Migra binds sem sgp_auto_cadastrar_os (padrão: False)
                    if "sgp_auto_cadastrar_os" not in b:
                        b["sgp_auto_cadastrar_os"] = False
                # Migra config sem papervines (v4 → v5)
                if "papervines" not in d:
                    d["papervines"] = dict(DEFAULT_CONFIG["papervines"])
                return d
            except Exception:
                pass
        self._write(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))

    def save(self):
        self._write(self.data)

    def _write(self, data):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def reload(self):
        self.data = self._load()

    def get_binds(self):
        return self.data.get("binds", [])

    def get_sgp(self):
        return self.data.get("sgp", {})

    def get_delay(self):
        return self.data.get("sgp", {}).get("delay_ms", 400) / 1000

    def get_papervines(self):
        default = DEFAULT_CONFIG["papervines"]
        return self.data.get("papervines", default)

    def update_papervines(self, **kwargs):
        if "papervines" not in self.data:
            self.data["papervines"] = dict(DEFAULT_CONFIG["papervines"])
        self.data["papervines"].update(kwargs)
        self.save()

    def update_bind(self, bind_id, **kwargs):
        for b in self.data["binds"]:
            if b["id"] == bind_id:
                b.update(kwargs)
                break
        self.save()

# ════════════════════════════════════════════════════════════
#  (OCR removido — leitura de número agora via HTML/Selenium)
# ════════════════════════════════════════════════════════════

# ════════════════════════════════════════════════════════════
#  SELENIUM SGP
# ════════════════════════════════════════════════════════════

class SGPSelenium:
    def __init__(self, cfg: ConfigManager):
        self.cfg       = cfg
        self.driver    = None
        self._link_buf = None
        self._link_ts  = 0

    # ── Conexão ─────────────────────────────────────────────

    def conectar(self) -> bool:
        if not SELENIUM_OK:
            messagebox.showerror("Selenium não instalado", "Execute instalar.bat")
            return False

        porta = self.cfg.get_sgp().get("debug_port", 9222)
        log.info(f"Conectando ao Chrome na porta {porta}...")
        try:
            opts = webdriver.ChromeOptions()
            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{porta}")
            self.driver = webdriver.Chrome(options=opts)
            log.info(f"Chrome conectado! Título: {self.driver.title}")
            return True
        except WebDriverException as e:
            msg = str(e)
            log.error(f"Erro ao conectar: {msg[:200]}")
            if "cannot connect" in msg.lower() or "connection refused" in msg.lower():
                messagebox.showerror(
                    "Chrome não encontrado",
                    f"Não foi possível conectar ao Chrome na porta {porta}.\n\n"
                    "Solução:\n"
                    "1. Feche todo o Chrome aberto\n"
                    "2. Execute  chrome_debug.bat\n"
                    "3. Acesse o SGP normalmente\n"
                    "4. Clique em Conectar ao Chrome novamente"
                )
            else:
                messagebox.showerror("Erro Selenium", msg[:400])
            return False
        except Exception as e:
            log.error(f"Erro inesperado: {e}")
            messagebox.showerror("Erro", str(e)[:400])
            return False

    def esta_conectado(self) -> bool:
        try:
            _ = self.driver.title
            return True
        except Exception:
            return False

    def _garantir_conexao(self) -> bool:
        if self.driver is None or not self.esta_conectado():
            return self.conectar()
        return True

    # ── Helpers ──────────────────────────────────────────────

    def _delay(self):
        time.sleep(self.cfg.get_delay())

    def _aguardar_clicavel(self, by, valor, timeout=8):
        return WebDriverWait(self.driver, timeout).until(
            EC.element_to_be_clickable((by, valor))
        )

    def _aguardar(self, by, valor, timeout=8):
        return WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, valor))
        )

    # ── Detecção de aba ativa (sem trocar abas visualmente) ───

    def _obter_tabs_cdp(self) -> list:
        """
        Consulta Chrome DevTools /json para obter TODAS as abas
        com seus IDs e URLs, sem trocar de aba.
        Os IDs do CDP coincidem com os window_handles do Selenium.
        """
        porta = self.cfg.get_sgp().get("debug_port", 9222)
        try:
            resp = urllib.request.urlopen(
                f"http://localhost:{porta}/json", timeout=3
            )
            tabs = json.loads(resp.read())
            pages = [t for t in tabs if t.get("type") == "page"]
            log.debug(f"  CDP: {len(pages)} abas tipo 'page'")
            return pages
        except Exception as e:
            log.debug(f"  CDP /json falhou: {e}")
            return []

    def _focar_aba_ocorrencia(self) -> bool:
        """
        Encontra a aba do formulário de ocorrência SEM trocar abas visualmente.
        Usa CDP /json para ler URLs de todas as abas de uma vez.
        Faz UM ÚNICO switch_to.window() direto para a aba certa.
        """
        try:
            handles = self.driver.window_handles
            log.debug(f"_focar_aba_ocorrencia: {len(handles)} aba(s)")

            # 1. Obtém todas as abas via CDP (sem switch visual)
            cdp_tabs = self._obter_tabs_cdp()

            if cdp_tabs:
                # A primeira entrada 'page' do CDP é a aba ativa do usuário
                aba_ativa_url = cdp_tabs[0].get("url", "") if cdp_tabs else ""
                log.debug(f"  Aba ativa do Chrome: {aba_ativa_url}")

                # Busca candidatas com 'ocorrencia' na URL
                candidatas = []
                for tab in cdp_tabs:
                    url = tab.get("url", "")
                    tab_id = tab.get("id", "")
                    if "ocorrencia" in url.lower():
                        candidatas.append((tab_id, url))

                # Prioridade: aba ativa > última candidata
                target_id = None
                target_url = None

                if "ocorrencia" in aba_ativa_url.lower():
                    target_id = cdp_tabs[0].get("id", "")
                    target_url = aba_ativa_url
                    log.info(f"  Aba ATIVA tem ocorrência: {target_url}")
                elif candidatas:
                    target_id, target_url = candidatas[-1]
                    log.info(f"  Usando última aba com ocorrência: {target_url}")

                if target_id:
                    # Tenta switch direto pelo ID do CDP (= window handle)
                    if target_id in handles:
                        self.driver.switch_to.window(target_id)
                        log.info(f"  Switch direto para handle {target_id[:12]}...")
                        return True

                    # Fallback: CDP ID não bate com handles — match por URL
                    log.debug("  CDP ID não encontrado nos handles, buscando por URL...")
                    for handle in handles:
                        self.driver.switch_to.window(handle)
                        if self.driver.current_url == target_url:
                            log.info(f"  Encontrado por URL match")
                            return True

                if not candidatas:
                    urls = [t.get("url", "")[:80] for t in cdp_tabs]
                    log.warning(f"  Nenhuma aba com 'ocorrencia'. URLs: {urls}")
                    messagebox.showerror(
                        "Formulário não encontrado",
                        "Nenhuma aba aberta contém o formulário de ocorrência.\n\n"
                        f"Abas encontradas ({len(cdp_tabs)}):\n"
                        + "\n".join(f"• {u}" for u in urls)
                        + "\n\nAbra o formulário no SGP e tente novamente."
                    )
                    return False

            # 2. Fallback total: CDP falhou — escaneia handles (com switch visual)
            log.warning("  CDP indisponível, fallback para scan de handles")
            abas_ocorrencia = []
            urls_encontradas = []

            for handle in handles:
                self.driver.switch_to.window(handle)
                url = self.driver.current_url
                urls_encontradas.append(url)
                if "ocorrencia" in url.lower():
                    abas_ocorrencia.append((handle, url))

            if abas_ocorrencia:
                handle, url = abas_ocorrencia[-1]
                self.driver.switch_to.window(handle)
                log.info(f"  Fallback: usando aba {url}")
                return True

            log.warning(f"  Nenhuma aba com 'ocorrencia'. URLs: {urls_encontradas}")
            messagebox.showerror(
                "Formulário não encontrado",
                "Nenhuma aba aberta contém o formulário de ocorrência.\n\n"
                + "\n".join(f"• {u[:80]}" for u in urls_encontradas)
                + "\n\nAbra o formulário no SGP e tente novamente."
            )
            return False

        except Exception as e:
            log.error(f"_focar_aba_ocorrencia erro: {e}", exc_info=True)
            messagebox.showerror("Erro ao trocar aba", str(e))
            return False

    def _selecionar_select2(self, container_id: str, filtro: str, nome: str) -> bool:
        """
        Preenche um dropdown Select2 via API jQuery (mais robusto que DOM).
        Extrai o ID do <select> real a partir do container_id.
        Fallback: tenta via DOM se jQuery não estiver disponível.
        """
        # container_id: "select2-id_tipo-container" → select_id: "id_tipo"
        select_id = container_id.replace("select2-", "").replace("-container", "")
        log.debug(f"_selecionar_select2: '{nome}' "
                  f"select_id='{select_id}' filtro='{filtro}'")
        try:
            # ── Método 1: jQuery Select2 API (funciona em classic e default) ──
            js_select2 = """
            var selectId = arguments[0];
            var filtro   = arguments[1].toLowerCase();
            var $sel     = jQuery('#' + selectId);

            if (!$sel.length) return 'NOT_FOUND';

            // Procura a primeira opção que contenha o filtro no texto
            var match = null;
            $sel.find('option').each(function() {
                if (!match && jQuery(this).text().toLowerCase().indexOf(filtro) !== -1) {
                    match = jQuery(this).val();
                }
            });

            if (match === null) return 'NO_MATCH';

            $sel.val(match).trigger('change');
            return 'OK:' + match;
            """
            resultado = self.driver.execute_script(js_select2, select_id, filtro)
            log.debug(f"  Select2 jQuery resultado: {resultado}")

            if resultado and resultado.startswith("OK:"):
                valor = resultado[3:]
                log.info(f"  Select2 '{nome}' → valor '{valor}' via jQuery API")
                self._delay()
                return True

            if resultado == "NOT_FOUND":
                log.warning(f"  Select #{select_id} não encontrado no DOM")
            elif resultado == "NO_MATCH":
                log.warning(f"  Nenhuma opção contém '{filtro}' no select #{select_id}")

            # ── Método 2: Fallback DOM (caso jQuery não funcione) ──
            log.debug(f"  Tentando fallback via DOM para '{nome}'")
            return self._selecionar_select2_dom(container_id, filtro, nome)

        except Exception as e:
            log.warning(f"  Select2 jQuery falhou para '{nome}': {e}")
            # Tenta fallback DOM
            return self._selecionar_select2_dom(container_id, filtro, nome)

    def _selecionar_select2_dom(self, container_id: str, filtro: str, nome: str) -> bool:
        """
        Fallback: abre Select2 via click DOM + ActionChains.
        Usado quando jQuery não está disponível.
        """
        log.debug(f"  _selecionar_select2_dom: '{nome}' container='{container_id}'")
        try:
            from selenium.webdriver.common.action_chains import ActionChains

            # 1. Localiza o .select2-selection pelo aria-labelledby
            css = f"[aria-labelledby='{container_id}']"
            selection = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, css))
            )

            # 2. Abre o dropdown via JS
            self.driver.execute_script("arguments[0].click();", selection)
            time.sleep(0.5)

            # 3. Tenta múltiplos seletores para o campo de busca
            #    (Select2 classic usa classes diferentes)
            seletores_busca = [
                ".select2-search__field",
                ".select2-search input",
                "input.select2-input",
                ".select2-drop .select2-input",
                ".select2-container--open input[type='search']",
            ]
            search = None
            for sel in seletores_busca:
                try:
                    search = WebDriverWait(self.driver, 2).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, sel))
                    )
                    log.debug(f"  Campo de busca encontrado via: {sel}")
                    break
                except TimeoutException:
                    continue

            if search is None:
                # Último recurso: digita direto no elemento ativo
                log.debug("  Nenhum campo de busca visível, digitando via JS")
                self.driver.execute_script("""
                    var evt = new Event('input', {bubbles: true});
                    var field = document.querySelector('.select2-search__field')
                              || document.querySelector('.select2-input');
                    if (field) {
                        field.value = arguments[0];
                        field.dispatchEvent(evt);
                    }
                """, filtro)
                time.sleep(0.5)
            else:
                ActionChains(self.driver).click(search).send_keys(filtro).perform()
                time.sleep(0.5)

            # 4. Clica na primeira opção
            opcao = WebDriverWait(self.driver, 6).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,
                     ".select2-results__option, .select2-result-selectable")
                )
            )
            texto_opcao = opcao.text[:60] if opcao.text else "(vazio)"
            self.driver.execute_script("arguments[0].click();", opcao)
            self._delay()
            log.info(f"  Select2 DOM '{nome}' → '{texto_opcao}'")
            return True

        except TimeoutException:
            log.error(f"Select2 DOM '{nome}' timeout")
            messagebox.showerror(
                "Campo não encontrado",
                f'"{nome}" não respondeu.\n\n'
                f'Verifique se o formulário de ocorrência está aberto no SGP\n'
                f'antes de pressionar a tecla.'
            )
            return False
        except Exception as e:
            log.error(f"Select2 DOM '{nome}' erro: {e}", exc_info=True)
            messagebox.showerror("Erro Select2", f'"{nome}": {e}')
            return False

    # ── Textarea ────────────────────────────────────────────

    def _preencher_textarea(self, field_id: str, texto: str, nome: str) -> bool:
        try:
            campo = self._aguardar_clicavel(By.ID, field_id)
            campo.click()
            campo.clear()
            campo.send_keys(texto)
            self._delay()
            return True
        except TimeoutException:
            messagebox.showerror(
                "Campo não encontrado",
                f'"{nome}" não foi localizado.\n'
                f'Verifique se o formulário está aberto.'
            )
            return False
        except Exception as e:
            messagebox.showerror("Erro", f'"{nome}": {e}')
            return False

    # ── Checkbox ────────────────────────────────────────────

    def _desmarcar_checkbox(self, field_id: str) -> bool:
        try:
            chk = self._aguardar(By.ID, field_id)
            if chk.is_selected():
                chk.click()
                time.sleep(0.2)
            return True
        except Exception as e:
            messagebox.showerror("Erro", f'Checkbox "Gerar OS?": {e}')
            return False

    # ── Botão ────────────────────────────────────────────────

    def _clicar_botao(self, field_id: str, nome: str) -> bool:
        try:
            btn = self._aguardar_clicavel(By.ID, field_id)
            btn.click()
            self._delay()
            return True
        except TimeoutException:
            messagebox.showerror(
                "Botão não encontrado",
                f'Botão "{nome}" não localizado.\nO formulário ainda está aberto?'
            )
            return False
        except Exception as e:
            messagebox.showerror("Erro", f'Botão "{nome}": {e}')
            return False

    # ── FLUXO: Ocorrência SGP (ULTRA-RÁPIDO) ─────────────────

    def _js_preencher_formulario(self) -> str:
        """Retorna o JS que preenche o formulário inteiro."""
        return """
        var tipoFiltro   = arguments[0];
        var origemFiltro = arguments[1].toLowerCase();
        var conteudo     = arguments[2];
        var desmarcarOS  = arguments[3];
        var autoCadastrar= arguments[4];
        var erros = [];

        // Detecção: estamos em página de ocorrência?
        if (typeof jQuery === 'undefined') return 'NO_JQUERY';
        if (!jQuery('#id_conteudo').length) return 'WRONG_PAGE';

        // Auto-detectar VMA → trocar sus por finan
        var host = window.location.hostname.toLowerCase();
        if (tipoFiltro === 'sus') {
            var vmaHosts = ['supersonic'];
            for (var i = 0; i < vmaHosts.length; i++) {
                if (host.indexOf(vmaHosts[i]) !== -1) {
                    tipoFiltro = 'finan';
                    break;
                }
            }
        }
        tipoFiltro = tipoFiltro.toLowerCase();

        // 1. Contrato
        try {
            var $cont = jQuery('#id_clientecontrato');
            if ($cont.length) {
                var opts = $cont.find('option').filter(function(){
                    return jQuery(this).val() !== '';
                });
                if (opts.length > 0) {
                    $cont.val(opts.first().val()).trigger('change');
                }
            }
        } catch(e) { erros.push('contrato:' + e.message); }

        // 2. Tipo de Ocorrência
        if (tipoFiltro) {
            try {
                var $tipo = jQuery('#id_tipo');
                if ($tipo.length) {
                    var match = null;
                    $tipo.find('option').each(function() {
                        if (!match && jQuery(this).text().toLowerCase().indexOf(tipoFiltro) !== -1) {
                            match = jQuery(this).val();
                        }
                    });
                    if (match !== null) $tipo.val(match).trigger('change');
                    else erros.push('tipo:NO_MATCH(' + tipoFiltro + ')');
                } else erros.push('tipo:NOT_FOUND');
            } catch(e) { erros.push('tipo:' + e.message); }
        }

        // 3. Origem
        try {
            var $orig = jQuery('#id_metodo');
            if ($orig.length) {
                var matchO = null;
                $orig.find('option').each(function() {
                    if (!matchO && jQuery(this).text().toLowerCase().indexOf(origemFiltro) !== -1) {
                        matchO = jQuery(this).val();
                    }
                });
                if (matchO !== null) $orig.val(matchO).trigger('change');
                else erros.push('origem:NO_MATCH(' + origemFiltro + ')');
            } else erros.push('origem:NOT_FOUND');
        } catch(e) { erros.push('origem:' + e.message); }

        // 4. Conteúdo
        try {
            jQuery('#id_conteudo').val(conteudo).trigger('change');
        } catch(e) { erros.push('conteudo:' + e.message); }

        // 5. Gerar OS
        if (desmarcarOS) {
            try {
                var chk = document.getElementById('id_os');
                if (chk && chk.checked) chk.click();
            } catch(e) {}
        }

        // 6. Cadastrar
        if (autoCadastrar) {
            try {
                var btn = document.getElementById('btacao');
                if (btn) btn.click();
            } catch(e) { erros.push('cadastrar:' + e.message); }
        }

        return erros.length === 0 ? 'OK' : 'ERROS:' + erros.join('|');
        """

    def executar_ocorrencia(self, bind: dict, msg_substituida: str):
        """
        Preenche o formulário via 1 chamada JS.
        Estratégia ultra-rápida:
          1. Tenta executar direto na aba atual (1 round-trip)
          2. Se falhar (página errada), aí sim faz detecção de aba
          3. Se auto_cadastrar_os=True → após cadastrar, preenche OS corretiva
        """
        t0 = time.time()
        log.info(f"=== executar_ocorrencia: '{bind.get('name')}' ===")

        if self.driver is None:
            if not self.conectar():
                return

        tipo_filtro    = bind.get("sgp_tipo_filtro",   "")
        origem_filtro  = bind.get("sgp_origem_filtro",  "whatsapp")
        desmarcar_os   = bind.get("sgp_desmarcar_os",   True)
        auto_cadastrar = bind.get("sgp_auto_cadastrar", False)
        auto_os        = bind.get("sgp_auto_cadastrar_os", False)

        # Se auto OS está ativo: força cadastrar=True e OS marcado
        if auto_os:
            auto_cadastrar = True
            desmarcar_os = False  # Precisa do checkbox "Gerar OS" MARCADO

        js = self._js_preencher_formulario()
        args = [tipo_filtro, origem_filtro, msg_substituida,
                desmarcar_os, auto_cadastrar]

        resultado = None

        # ── Tentativa 1: executa direto (sem trocar aba) ──
        try:
            resultado = self.driver.execute_script(js, *args)
            elapsed = (time.time() - t0) * 1000

            if resultado == "OK" or (resultado and resultado.startswith("ERROS:")):
                log.info(f"  Direto em {elapsed:.0f}ms → {resultado}")
            else:
                resultado = None  # Precisa tentar outra aba
                log.debug(f"  Aba atual não é formulário, buscando...")

        except Exception as e:
            log.debug(f"  Execução direta falhou: {e}")
            if not self.esta_conectado():
                if not self.conectar():
                    return

        # ── Tentativa 2: detecta aba e tenta de novo ──
        if resultado is None:
            if not self._focar_aba_ocorrencia():
                return
            try:
                resultado = self.driver.execute_script(js, *args)
                elapsed = (time.time() - t0) * 1000
                log.info(f"  Com tab switch em {elapsed:.0f}ms → {resultado}")
            except Exception as e:
                log.error(f"  Falha no preenchimento: {e}", exc_info=True)
                messagebox.showerror("Erro SGP", f"Falha ao preencher: {e}")
                return

        # ── Etapa 3: Preencher formulário de OS (se auto_os ativo) ──
        if auto_os and resultado and (resultado == "OK" or resultado.startswith("ERROS:")):
            log.info("  Auto OS ativo → aguardando formulário de OS...")
            self._preencher_os_corretiva()

    def _preencher_os_corretiva(self):
        """
        Após cadastrar ocorrência com Gerar OS marcado, o SGP redireciona
        para o formulário de OS. Este método:
          1. Aguarda a página de OS carregar
          2. Seleciona 'corretiva' no campo Motivo (Select2 #id_motivoos)
          3. Clica em Cadastrar
        """
        try:
            # Aguarda a página de OS carregar (campo motivo aparece)
            log.debug("  Aguardando campo Motivo na página de OS...")
            time.sleep(1.5)  # Espera navegação do formulário

            js_os = """
            if (typeof jQuery === 'undefined') return 'NO_JQUERY';

            // Verifica se o campo motivo existe (indica que estamos na pág de OS)
            var $motivo = jQuery('#id_motivoos');
            if (!$motivo.length) return 'NO_MOTIVO';

            // Seleciona "corretiva" no dropdown de motivo
            var match = null;
            $motivo.find('option').each(function() {
                if (!match && jQuery(this).text().toLowerCase().indexOf('corretiva') !== -1) {
                    match = jQuery(this).val();
                }
            });
            if (match === null) return 'NO_MATCH_CORRETIVA';

            $motivo.val(match).trigger('change');

            // Clica em Cadastrar
            var btn = document.getElementById('btacao');
            if (btn) btn.click();

            return 'OS_OK';
            """

            # Tenta executar na aba atual (deve ser a OS após redirect)
            resultado = self.driver.execute_script(js_os)
            log.info(f"  OS corretiva resultado: {resultado}")

            if resultado == "NO_MOTIVO":
                # Talvez a página ainda não carregou, tenta de novo
                log.debug("  Campo motivo não encontrado, aguardando mais...")
                time.sleep(2)
                resultado = self.driver.execute_script(js_os)
                log.info(f"  OS corretiva retry: {resultado}")

            if resultado == "NO_MATCH_CORRETIVA":
                log.warning("  Opção 'corretiva' não encontrada no dropdown de motivo")
                messagebox.showwarning(
                    "Motivo não encontrado",
                    "A opção 'corretiva' não foi encontrada no campo Motivo.\n"
                    "Verifique o formulário de OS manualmente."
                )

        except Exception as e:
            log.error(f"  Erro ao preencher OS corretiva: {e}", exc_info=True)
            messagebox.showerror("Erro OS", f"Falha ao preencher formulário de OS:\n{e}")

    # ── LEITURA DE NÚMERO VIA HTML (substitui OCR) ───────────

    def capturar_numero_html(self) -> str:
        """
        Lê TODOS os números de telefone do <select id="id_protocolo_sms">.
        Cada option pode ter texto tipo "(11) 98236-3966 - Celular Pessoal"
        ou "(11) 99999-1234 - WhatsApp". Extrai o telefone de TODAS as options.
        """
        if self.driver is None:
            if not self.conectar():
                return ""

        js = """
        var sel = document.getElementById('id_protocolo_sms');
        if (!sel) return '';
        var textos = [];
        for (var i = 0; i < sel.options.length; i++) {
            var t = sel.options[i].text.trim();
            if (t) textos.push(t);
        }
        return textos.join('|||');
        """

        def _extrair_todos(raw):
            """Extrai telefone de cada option, retorna separados por quebra de linha."""
            if not raw:
                return ""
            numeros = []
            for texto in raw.split('|||'):
                texto = texto.strip()
                if not texto:
                    continue
                # Formato: "(11) 9522-29558 - Celular Pessoal"
                # Pega tudo antes do último " - " (que é o rótulo)
                if " - " in texto:
                    numero = texto.rsplit(" - ", 1)[0].strip()
                else:
                    numero = texto
                if numero:
                    numeros.append(numero)
            return '\n'.join(numeros)

        # Tentativa 1: executa na aba atual
        try:
            raw = self.driver.execute_script(js)
            if raw:
                resultado = _extrair_todos(raw)
                log.info(f"  Números via HTML (aba atual): '{resultado}'")
                return resultado
        except Exception as e:
            log.debug(f"  Leitura HTML falhou na aba atual: {e}")

        # Tentativa 2: foca aba de ocorrência e tenta de novo
        if self._focar_aba_ocorrencia():
            try:
                raw = self.driver.execute_script(js)
                if raw:
                    resultado = _extrair_todos(raw)
                    log.info(f"  Números via HTML (após switch): '{resultado}'")
                    return resultado
            except Exception as e:
                log.error(f"  Leitura HTML falhou após switch: {e}")

        log.warning("  Nenhum número encontrado em id_protocolo_sms")
        return ""

    # ── FLUXO: Enviar texto no chat ──────────────────────────

    def enviar_chat(self, texto: str):
        """
        Cola texto inteiro na janela ativa de uma vez só (instantâneo).
        Usa Ctrl+V com o texto completo no clipboard.
        WhatsApp Web preserva quebras de linha ao colar do clipboard.
        """
        if not PYAUTOGUI_OK:
            return
        clip_ant = pyperclip.paste()
        pyperclip.copy(texto)
        time.sleep(0.05)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.05)
        pyperclip.copy(clip_ant)

    # ── FLUXO: Link pagamento ────────────────────────────────

    def executar_link(self, bind: dict):
        if not PYAUTOGUI_OK:
            return
        agora = time.time()

        if self._link_buf is None:
            clip_ant = pyperclip.paste()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.15)
            link = pyperclip.paste()
            pyperclip.copy(clip_ant)

            if not re.match(r'https?://', link or '', re.IGNORECASE):
                return

            self._link_buf = link
            self._link_ts  = agora
            return

        if agora - self._link_ts > 10:
            self._link_buf = None
            return

        msg = bind.get("message", "{link}").replace("{link}", self._link_buf)
        self.enviar_chat(msg)
        self._link_buf = None

# ════════════════════════════════════════════════════════════
#  EXECUTOR DE BIND — roda todas as ações em sequência
# ════════════════════════════════════════════════════════════

class BindExecutor:
    """
    Centraliza a execução de uma bind com múltiplos tipos de ação.
    Ordem de execução:
      1. text_ocr  → captura número via HTML e envia no chat
      2. text      → envia texto fixo no chat
      3. sgp_ocorrencia → preenche formulário no SGP
      4. link_pagamento → fluxo de 2 etapas do link
    """

    def __init__(self, sgp: SGPSelenium):
        self.sgp = sgp

    def executar(self, bind: dict):
        tipos  = bind.get("types", ["text"])
        msg_template = bind.get("message", "")
        mes    = str(datetime.now().month)  # "3" para Março, "12" para Dezembro

        log.info(f"── BindExecutor: '{bind.get('name')}' "
                 f"tipos={tipos} key={bind.get('key')} ──")

        try:
            # Captura número via HTML (substituiu OCR) se necessário
            numero = ""
            if "text_ocr" in tipos:
                numero = self.sgp.capturar_numero_html()
                log.debug(f"  HTML resultado: '{numero[:40]}'" if numero else "  HTML: vazio")

            # Monta mensagem substituindo variáveis
            msg = msg_template
            msg = msg.replace("{mes}",    mes)
            msg = msg.replace("{numero}", numero)
            # {link} é tratado internamente pelo executar_link

            # Executa cada tipo na ordem
            if "text_ocr" in tipos:
                log.debug("  Executando: text_ocr → enviar_chat")
                self.sgp.enviar_chat(msg)

            if "text" in tipos and "text_ocr" not in tipos:
                log.debug("  Executando: text → enviar_chat")
                self.sgp.enviar_chat(msg)

            if "sgp_ocorrencia" in tipos:
                log.debug("  Executando: sgp_ocorrencia")
                self.sgp.executar_ocorrencia(bind, msg)

            if "link_pagamento" in tipos:
                log.debug("  Executando: link_pagamento")
                self.sgp.executar_link(bind)

            log.info(f"── BindExecutor: '{bind.get('name')}' concluído ──")

        except Exception as e:
            log.error(f"ERRO na bind '{bind.get('name')}': {e}", exc_info=True)
            try:
                messagebox.showerror(
                    "Erro na execução",
                    f"Bind: {bind.get('name')}\n"
                    f"Erro: {type(e).__name__}: {e}"
                )
            except Exception:
                pass  # UI pode não estar disponível

# ════════════════════════════════════════════════════════════
#  HOTKEY MANAGER
# ════════════════════════════════════════════════════════════

class HotkeyManager:
    def __init__(self):
        self._registradas = []

    def registrar(self, key: str, callback):
        if not KEYBOARD_OK:
            return
        try:
            keyboard.add_hotkey(key, callback, suppress=True)
            self._registradas.append(key)
        except Exception as e:
            print(f"[Hotkey] Erro ao registrar {key}: {e}")

    def limpar(self):
        if not KEYBOARD_OK:
            return
        for key in self._registradas:
            try:
                keyboard.remove_hotkey(key)
            except Exception:
                pass
        self._registradas.clear()

# ════════════════════════════════════════════════════════════
#  PAPERVINES AUTOMATION
# ════════════════════════════════════════════════════════════

class PapervinesAutomation:
    """
    Automação do Papervines: loop que processa fila de "Novos".
    Para cada cliente:
      1. Clica na aba/botão "Novos"
      2. Clica no primeiro cliente da lista
      3. Clica em "Iniciar"
      4. Digita saudação na caixa de texto
      5. Envia (clica no botão enviar)
      6. Repete até a fila zerar
    """

    def __init__(self, sgp: SGPSelenium, cfg: ConfigManager):
        self.sgp = sgp
        self.cfg = cfg
        self._rodando = False
        self._parar = False
        self._status_callback = None  # UI atualiza status

    @property
    def driver(self):
        return self.sgp.driver

    def esta_rodando(self):
        return self._rodando

    def parar(self):
        self._parar = True

    def _status(self, msg):
        log.info(f"[Papervines] {msg}")
        if self._status_callback:
            try:
                self._status_callback(msg)
            except Exception:
                pass

    def _focar_aba_papervines(self) -> bool:
        """Encontra e foca a aba do Papervines no Chrome."""
        try:
            cdp_tabs = self.sgp._obter_tabs_cdp()
            handles = self.driver.window_handles

            for tab in cdp_tabs:
                url = tab.get("url", "").lower()
                if "chat.papervines.digital" in url or "papervines" in url:
                    tab_id = tab.get("id", "")
                    if tab_id in handles:
                        self.driver.switch_to.window(tab_id)
                        log.info(f"  Papervines: switch para {url[:60]}")
                        return True
                    # Fallback por URL
                    for handle in handles:
                        self.driver.switch_to.window(handle)
                        if "papervines" in self.driver.current_url.lower():
                            return True

            # Fallback: scan todos os handles
            for handle in handles:
                self.driver.switch_to.window(handle)
                if "papervines" in self.driver.current_url.lower():
                    return True

            self._status("Aba do Papervines não encontrada!")
            return False
        except Exception as e:
            log.error(f"  Papervines focar aba erro: {e}")
            return False

    def executar_loop(self, saudacao: str, delay_ms: int):
        """
        Loop principal: processa todos os clientes na fila Novos.
        Executa via JavaScript no DOM do Papervines.
        """
        if self._rodando:
            self._status("Já está rodando!")
            return

        self._rodando = True
        self._parar = False
        atendidos = 0

        try:
            if self.driver is None:
                if not self.sgp.conectar():
                    self._status("Falha ao conectar ao Chrome")
                    return

            if not self._focar_aba_papervines():
                return

            delay_s = delay_ms / 1000
            self._status("Iniciando loop de atendimento...")

            while not self._parar:
                # 1. Verifica se há clientes na fila "Novos"
                count = self._contar_novos()
                if count == 0:
                    self._status(f"Fila zerada! {atendidos} clientes atendidos.")
                    break

                self._status(f"Fila: {count} | Atendidos: {atendidos}")

                # 2. Clica na aba "Novos" (garante que está na aba certa)
                if not self._clicar_novos():
                    self._status("Erro ao clicar em 'Novos'")
                    break

                time.sleep(0.5)

                # 3. Clica no primeiro cliente da lista
                if not self._clicar_primeiro_cliente():
                    self._status(f"Sem mais clientes. {atendidos} atendidos.")
                    break

                time.sleep(0.8)

                # 4. Clica em "Iniciar"
                if not self._clicar_iniciar():
                    self._status("Botão 'Iniciar' não encontrado")
                    time.sleep(1)
                    continue

                time.sleep(2)  # Espera chat carregar após Iniciar

                # 5. Digita saudação e envia (com retry)
                if self._enviar_saudacao(saudacao):
                    atendidos += 1
                    self._status(f"Enviado! Total: {atendidos}")
                else:
                    self._status("Erro ao enviar saudação")

                time.sleep(delay_s)

        except Exception as e:
            log.error(f"[Papervines] Loop erro: {e}", exc_info=True)
            self._status(f"Erro: {e}")
        finally:
            self._rodando = False
            self._parar = False
            if atendidos > 0:
                self._status(f"Finalizado! {atendidos} clientes atendidos.")

    def _contar_novos(self) -> int:
        """Conta quantos clientes estão na fila Novos via badge vermelho."""
        try:
            js = """
            // Badge vermelho dentro do botão Novos: data-cy="button-new-sessions"
            var btn = document.querySelector('[data-cy="button-new-sessions"]');
            if (!btn) return 0;
            var badge = btn.querySelector('.bg-red-600, .rounded-full.text-white');
            if (badge) {
                var n = parseInt(badge.textContent.trim());
                if (!isNaN(n)) return n;
            }
            // Fallback: conta itens na lista visível
            var items = document.querySelectorAll('[data-cy="list-item-session"]');
            return items.length;
            """
            result = self.driver.execute_script(js)
            return int(result) if result else 0
        except Exception:
            return 0

    def _clicar_novos(self) -> bool:
        """Clica no botão 'Novos' (data-cy="button-new-sessions")."""
        try:
            js = """
            var el = document.querySelector('[data-cy="button-new-sessions"]');
            if (el) { el.click(); return 'OK'; }
            // Fallback: busca por texto
            var spans = document.querySelectorAll('span');
            for (var i = 0; i < spans.length; i++) {
                var txt = spans[i].textContent.trim();
                if (txt === 'Novos' || txt === 'Novo') {
                    spans[i].closest('[data-cy], .cursor-pointer, span.flex').click();
                    return 'OK_TEXT';
                }
            }
            return 'NOT_FOUND';
            """
            result = self.driver.execute_script(js)
            return result in ("OK", "OK_TEXT")
        except Exception as e:
            log.error(f"  Papervines clicar Novos erro: {e}")
            return False

    def _clicar_primeiro_cliente(self) -> bool:
        """Clica no primeiro cliente da lista (data-cy="list-item-session")."""
        try:
            js = """
            var el = document.querySelector('[data-cy="list-item-session"]');
            if (el) { el.click(); return 'OK'; }
            return 'NOT_FOUND';
            """
            result = self.driver.execute_script(js)
            return result == "OK"
        except Exception as e:
            log.error(f"  Papervines clicar cliente erro: {e}")
            return False

    def _clicar_iniciar(self) -> bool:
        """Clica no botão 'Iniciar' (data-cy="button-session-start")."""
        try:
            js = """
            // Seletor exato: h-button com data-cy, depois o button interno
            var hbtn = document.querySelector('[data-cy="button-session-start"]');
            if (hbtn) {
                var btn = hbtn.querySelector('button') || hbtn;
                btn.click();
                return 'OK';
            }
            // Fallback: busca por texto "Iniciar" em qualquer button
            var btns = document.querySelectorAll('button');
            for (var i = 0; i < btns.length; i++) {
                if (btns[i].textContent.trim().toLowerCase().indexOf('iniciar') !== -1) {
                    btns[i].click();
                    return 'OK_TEXT';
                }
            }
            return 'NOT_FOUND';
            """
            result = self.driver.execute_script(js)
            return result in ("OK", "OK_TEXT")
        except Exception as e:
            log.error(f"  Papervines clicar Iniciar erro: {e}")
            return False

    def _enviar_saudacao(self, texto: str) -> bool:
        """
        Digita saudação no textarea do Papervines e envia via Enter.
        Tem retry pois o chat pode demorar para carregar após Iniciar.
        """
        js_type = """
        var texto = arguments[0];

        // Tenta múltiplos seletores para o textarea
        var ta = document.querySelector('[data-cy="message-input"] textarea')
              || document.querySelector('textarea.mat-input-element')
              || document.querySelector('mat-form-field textarea')
              || document.querySelector('textarea[rows]')
              || document.querySelector('.chat-footer textarea')
              || document.querySelector('.reply-box textarea');
        if (!ta) return 'NO_INPUT';

        // Foca e seta valor
        ta.focus();
        ta.value = texto;

        // Dispara eventos para Angular detectar a mudança
        ta.dispatchEvent(new Event('input', {bubbles: true}));
        ta.dispatchEvent(new Event('change', {bubbles: true}));

        return 'TYPED';
        """

        js_send = """
        var ta = document.querySelector('[data-cy="message-input"] textarea')
              || document.querySelector('textarea.mat-input-element')
              || document.querySelector('mat-form-field textarea')
              || document.querySelector('textarea[rows]');
        if (!ta) return 'NO_INPUT';

        // Simula Enter keydown + keyup (como o Angular espera)
        ta.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Enter', code: 'Enter', keyCode: 13, which: 13,
            bubbles: true, cancelable: true
        }));
        ta.dispatchEvent(new KeyboardEvent('keyup', {
            key: 'Enter', code: 'Enter', keyCode: 13, which: 13,
            bubbles: true
        }));
        return 'SENT';
        """

        # Retry até 5 tentativas (textarea pode demorar a carregar)
        for tentativa in range(5):
            try:
                result = self.driver.execute_script(js_type, texto)
                if result == "TYPED":
                    time.sleep(0.3)
                    result2 = self.driver.execute_script(js_send)
                    if result2 == "SENT":
                        return True
                    log.debug(f"  Papervines envio tentativa {tentativa+1}: {result2}")
                else:
                    log.debug(f"  Papervines textarea tentativa {tentativa+1}: {result}")

                # Espera mais antes de tentar de novo
                time.sleep(1)

            except Exception as e:
                log.debug(f"  Papervines saudação tentativa {tentativa+1} erro: {e}")
                time.sleep(1)

        log.warning("  Papervines: campo de texto não encontrado após 5 tentativas")
        return False

    # ── TRANSFERÊNCIA EM MASSA ───────────────────────────────

    def executar_transferencia(self, departamento: str, delay_ms: int):
        """
        Loop de transferência: pega cada cliente de Novos e transfere
        para o departamento especificado.
        Fluxo: Novos → Cliente → Transferir → Selecionar dept → Confirmar → Repetir
        """
        if self._rodando:
            self._status("Já está rodando!")
            return

        self._rodando = True
        self._parar = False
        transferidos = 0

        try:
            if self.driver is None:
                if not self.sgp.conectar():
                    self._status("Falha ao conectar ao Chrome")
                    return

            if not self._focar_aba_papervines():
                return

            delay_s = delay_ms / 1000
            self._status(f"Transferindo para: {departamento}")

            # 1. Clica em Novos (só na primeira vez)
            if not self._clicar_novos():
                self._status("Erro ao clicar em 'Novos'")
                return
            time.sleep(0.8)

            while not self._parar:
                count = self._contar_novos()
                if count == 0:
                    self._status(f"Fila zerada! {transferidos} transferidos.")
                    break

                self._status(f"Fila: {count} | Transferidos: {transferidos}")

                # 2. Clica no primeiro cliente
                if not self._clicar_primeiro_cliente():
                    self._status(f"Sem mais clientes. {transferidos} transferidos.")
                    break

                time.sleep(0.8)

                # 3. Clica no botão Transferir (abre modal)
                if not self._clicar_btn_transferir():
                    self._status("Botão 'Transferir' não encontrado")
                    time.sleep(1)
                    continue

                time.sleep(1)

                # 4. Seleciona o departamento na lista
                if not self._selecionar_departamento(departamento):
                    self._status(f"Departamento '{departamento}' não encontrado")
                    time.sleep(1)
                    continue

                time.sleep(0.5)

                # 5. Clica no botão confirmar transferência
                if not self._confirmar_transferencia():
                    self._status("Erro ao confirmar transferência")
                    time.sleep(1)
                    continue

                transferidos += 1
                self._status(f"Transferido! Total: {transferidos}")
                time.sleep(delay_s)

        except Exception as e:
            log.error(f"[Papervines] Transferência erro: {e}", exc_info=True)
            self._status(f"Erro: {e}")
        finally:
            self._rodando = False
            self._parar = False
            if transferidos > 0:
                self._status(f"Finalizado! {transferidos} clientes transferidos.")

    def _clicar_btn_transferir(self) -> bool:
        """Clica no botão 'Transferir' do cliente (data-cy="button-session-transfer")."""
        try:
            js = """
            var hbtn = document.querySelector('[data-cy="button-session-transfer"]');
            if (hbtn) {
                var btn = hbtn.querySelector('button') || hbtn;
                btn.click();
                return 'OK';
            }
            // Fallback: busca por texto
            var btns = document.querySelectorAll('button');
            for (var i = 0; i < btns.length; i++) {
                if (btns[i].textContent.trim().toLowerCase().indexOf('transferir') !== -1) {
                    btns[i].click();
                    return 'OK_TEXT';
                }
            }
            return 'NOT_FOUND';
            """
            result = self.driver.execute_script(js)
            return result in ("OK", "OK_TEXT")
        except Exception as e:
            log.error(f"  Papervines clicar Transferir erro: {e}")
            return False

    def _selecionar_departamento(self, nome_dept: str) -> bool:
        """
        Seleciona um departamento na lista de transferência pelo nome.
        Primeiro garante que está na aba 'Para uma equipe', depois clica
        no departamento cujo texto contém nome_dept.
        """
        nome_lower = nome_dept.lower().strip()

        # Retry: a lista pode demorar pra carregar
        for tentativa in range(5):
            try:
                js = """
                var nomeDept = arguments[0].toLowerCase();

                // Garante aba "Para uma equipe" ativa
                var tabs = document.querySelectorAll('[role="tab"], .mat-tab-label, button');
                for (var i = 0; i < tabs.length; i++) {
                    var txt = (tabs[i].textContent || '').trim().toLowerCase();
                    if (txt.indexOf('equipe') !== -1 || txt.indexOf('team') !== -1) {
                        tabs[i].click();
                        break;
                    }
                }

                // Busca departamento na lista
                var items = document.querySelectorAll('[data-cy="department-list-name"]');
                for (var j = 0; j < items.length; j++) {
                    var itemText = (items[j].textContent || '').trim().toLowerCase();
                    if (itemText.indexOf(nomeDept) !== -1) {
                        items[j].click();
                        return 'OK';
                    }
                }

                // Fallback: busca em qualquer item de lista
                var divs = document.querySelectorAll('.cursor-pointer');
                for (var k = 0; k < divs.length; k++) {
                    var dText = (divs[k].textContent || '').trim().toLowerCase();
                    if (dText.indexOf(nomeDept) !== -1 && dText.length < 100) {
                        divs[k].click();
                        return 'OK_FALLBACK';
                    }
                }

                return 'NOT_FOUND';
                """
                result = self.driver.execute_script(js, nome_dept)
                if result in ("OK", "OK_FALLBACK"):
                    log.info(f"  Departamento '{nome_dept}' selecionado ({result})")
                    return True

                log.debug(f"  Departamento tentativa {tentativa+1}: {result}")
                time.sleep(1)

            except Exception as e:
                log.debug(f"  Departamento tentativa {tentativa+1} erro: {e}")
                time.sleep(1)

        log.warning(f"  Departamento '{nome_dept}' não encontrado após 5 tentativas")
        return False

    def _confirmar_transferencia(self) -> bool:
        """Clica no botão confirmar transferência (data-cy="button-next-department")."""
        for tentativa in range(3):
            try:
                js = """
                // Botão confirmar: data-cy="button-next-department"
                var hbtn = document.querySelector('[data-cy="button-next-department"]');
                if (hbtn) {
                    var btn = hbtn.querySelector('button') || hbtn;
                    btn.click();
                    return 'OK';
                }
                // Fallback: último botão "Transferir" visível (o do modal)
                var btns = document.querySelectorAll('button');
                var last = null;
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].textContent.trim().toLowerCase() === 'transferir'
                        && btns[i].offsetParent !== null) {
                        last = btns[i];
                    }
                }
                if (last) { last.click(); return 'OK_LAST'; }
                return 'NOT_FOUND';
                """
                result = self.driver.execute_script(js)
                if result in ("OK", "OK_LAST"):
                    return True
                log.debug(f"  Confirmar transferência tentativa {tentativa+1}: {result}")
                time.sleep(0.8)
            except Exception as e:
                log.debug(f"  Confirmar transferência tentativa {tentativa+1} erro: {e}")
                time.sleep(0.8)
        return False

# ════════════════════════════════════════════════════════════
#  UI UTILS
# ════════════════════════════════════════════════════════════

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def _rgb_to_hex(r, g, b):
    return f"#{int(r):02x}{int(g):02x}{int(b):02x}"

def _lerp_color(c1, c2, t):
    """Interpola entre duas cores hex. t=0 retorna c1, t=1 retorna c2."""
    r1, g1, b1 = _hex_to_rgb(c1)
    r2, g2, b2 = _hex_to_rgb(c2)
    return _rgb_to_hex(
        r1 + (r2 - r1) * t,
        g1 + (g2 - g1) * t,
        b1 + (b2 - b1) * t,
    )

def _font(size=12, bold=False, theme=None):
    """Retorna fonte baseada no tema ativo."""
    t = theme or THEME
    w = "bold" if bold else "normal"
    return ctk.CTkFont(family=t["font"], size=size, weight=w)

# ════════════════════════════════════════════════════════════
#  UI: JANELA PRINCIPAL — SGPilot
# ════════════════════════════════════════════════════════════

class FloatingApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.cfg        = ConfigManager()
        self.sgp        = SGPSelenium(self.cfg)
        self.executor   = BindExecutor(self.sgp)
        self.papervines = PapervinesAutomation(self.sgp, self.cfg)
        self.hotkeys    = HotkeyManager()
        self._tema_ativo = "sgp"
        self._animating  = False

        ctk.set_appearance_mode("dark")

        self._setup_janela()
        self._construir_ui()
        self._registrar_hotkeys()
        self._loop_topo()

    @property
    def T(self):
        return THEME_SGP if self._tema_ativo == "sgp" else THEME_PV

    def _setup_janela(self):
        self.title("SGPilot")
        self.geometry("320x600+60+180")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 1.0)
        self.configure(fg_color=THEME_SGP["bg"])
        self.protocol("WM_DELETE_WINDOW", self._fechar)

        if LOGO_ICON.exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(LOGO_ICON).resize((32, 32))
                self._icon = ImageTk.PhotoImage(img)
                self.iconphoto(False, self._icon)
            except Exception:
                try:
                    icon = tk.PhotoImage(file=str(LOGO_ICON))
                    self.iconphoto(False, icon)
                    self._icon = icon
                except Exception:
                    pass

    # ══════════════════════════════════════════════════════════
    #  CONSTRUÇÃO DA UI
    # ══════════════════════════════════════════════════════════

    def _construir_ui(self):
        T = THEME_SGP

        # ── Header: logo3 centralizado ──
        self.header = ctk.CTkFrame(self, height=48, corner_radius=0, fg_color=T["bg_header"])
        self.header.pack(fill="x")
        self.header.pack_propagate(False)

        self._header_logo = None
        logo_path = LOGO3 if LOGO3.exists() else LOGO_ICON
        if logo_path.exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                # Escala proporcional: altura 32px
                ratio = 32 / img.height
                new_w = int(img.width * ratio)
                img = img.resize((new_w, 32), Image.LANCZOS)
                self._header_logo = ImageTk.PhotoImage(img)
                ctk.CTkLabel(
                    self.header, text="", image=self._header_logo
                ).pack(side="left", padx=(14, 0), pady=8)
            except Exception:
                # Fallback: texto
                ctk.CTkLabel(self.header, text="SGPilot",
                    font=ctk.CTkFont(family="Roboto", size=14, weight="bold"),
                    text_color=T["text"]).pack(side="left", padx=14, pady=10)

        self.btn_config = ctk.CTkButton(
            self.header, text="⚙", width=32, height=28,
            command=self._abrir_config,
            fg_color="transparent", hover_color=T["bg_card_hover"],
            font=ctk.CTkFont(size=14), text_color=T["text_dim"]
        )
        self.btn_config.pack(side="right", padx=6, pady=4)

        # ── Barra de abas (logo abaixo, sem linha separadora) ──
        self.tab_bar = ctk.CTkFrame(self, height=42, corner_radius=0, fg_color=T["bg"])
        self.tab_bar.pack(fill="x")
        self.tab_bar.pack_propagate(False)

        inner = ctk.CTkFrame(self.tab_bar, fg_color="transparent")
        inner.pack(expand=True, pady=6)

        # SGP tab: #424242, texto #cacacc (ativo = com borda inferior accent)
        self.btn_tab_sgp = ctk.CTkButton(
            inner, text="SGP", height=28, width=130,
            command=lambda: self._trocar_tema("sgp"),
            corner_radius=4,
            fg_color=T["tab_active"], hover_color=T["bg_card_hover"],
            text_color=T["text"],
            font=ctk.CTkFont(family="Roboto", size=11, weight="bold"),
            border_width=2, border_color=T["accent"]
        )
        self.btn_tab_sgp.pack(side="left", padx=4)

        self.btn_tab_pv = ctk.CTkButton(
            inner, text="Papervines", height=28, width=130,
            command=lambda: self._trocar_tema("pv"),
            corner_radius=4,
            fg_color=T["bg_card"], hover_color=T["bg_card_hover"],
            text_color=T["text_dim"],
            font=ctk.CTkFont(family="Roboto", size=11, weight="bold"),
            border_width=1, border_color=T["border"]
        )
        self.btn_tab_pv.pack(side="left", padx=4)

        # ── Botão conectar: verde #28a745, fonte #fffbe7 ──
        self.connect_frame = ctk.CTkFrame(self, fg_color=T["bg"], corner_radius=0)
        self.connect_frame.pack(fill="x", padx=10, pady=(6, 0))

        self.btn_conectar = ctk.CTkButton(
            self.connect_frame, text="◇  Conectar ao Chrome", height=32,
            command=self._conectar_chrome, corner_radius=4,
            fg_color="#28a745", hover_color="#2fbd4f",
            text_color="#fffbe7",
            font=ctk.CTkFont(family="Roboto", size=11, weight="bold"),
        )
        self.btn_conectar.pack(fill="x")

        self.status_lbl = ctk.CTkLabel(
            self.connect_frame, text="○  Desconectado",
            font=ctk.CTkFont(family="Roboto", size=9), text_color=T["text_dim"]
        )
        self.status_lbl.pack(anchor="w", padx=4, pady=(2, 0))

        # ── Container de conteúdo ──
        self.content = ctk.CTkFrame(self, fg_color=T["bg"], corner_radius=0)
        self.content.pack(fill="both", expand=True)

        self.painel_sgp = ctk.CTkFrame(self.content, fg_color=T["bg"], corner_radius=0)
        self.painel_pv  = ctk.CTkFrame(self.content, fg_color=THEME_PV["bg"], corner_radius=0)

        self._construir_aba_sgp()
        self._construir_aba_papervines()
        self.painel_sgp.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════
    #  TROCA DE TEMA (fade suave)
    # ══════════════════════════════════════════════════════════

    def _trocar_tema(self, novo):
        if novo == self._tema_ativo or self._animating:
            return
        self._animating = True
        antigo = self._tema_ativo
        self._tema_ativo = novo

        T_old = THEME_SGP if antigo == "sgp" else THEME_PV
        T_new = THEME_SGP if novo == "sgp" else THEME_PV

        # Fase 1: Fade out (esmaecer o painel atual)
        steps_out = 8
        steps_in  = 8

        def fade_out(i):
            t = i / steps_out
            try:
                # Interpola cores do chrome/header junto
                bg = _lerp_color(T_old["bg"], T_new["bg"], t)
                bh = _lerp_color(T_old["bg_header"], T_new["bg_header"], t)
                ac = _lerp_color(T_old["accent"], T_new["accent"], t)
                # Aplica opacidade diminuindo no app
                alpha = 1.0 - (t * 0.3)  # 1.0 → 0.7
                self.attributes("-alpha", alpha)
                self.configure(fg_color=bg)
                self.header.configure(fg_color=bh)
                self.tab_bar.configure(fg_color=bg)
                self.connect_frame.configure(fg_color=bg)
                self.content.configure(fg_color=bg)
            except Exception:
                pass
            if i < steps_out:
                self.after(15, lambda: fade_out(i + 1))
            else:
                # Troca de modo e painel no meio da transição
                if novo == "pv":
                    ctk.set_appearance_mode("light")
                else:
                    ctk.set_appearance_mode("dark")
                if antigo == "sgp":
                    self.painel_sgp.pack_forget()
                    self.painel_pv.pack(fill="both", expand=True)
                else:
                    self.painel_pv.pack_forget()
                    self.painel_sgp.pack(fill="both", expand=True)
                fade_in(0)

        def fade_in(i):
            t = i / steps_in
            try:
                alpha = 0.7 + (t * 0.3)  # 0.7 → 1.0
                self.attributes("-alpha", alpha)
            except Exception:
                pass
            if i < steps_in:
                self.after(15, lambda: fade_in(i + 1))
            else:
                self.attributes("-alpha", 1.0)
                self._finalizar_troca(novo)

        fade_out(0)

    def _finalizar_troca(self, tema):
        T = THEME_SGP if tema == "sgp" else THEME_PV

        self.btn_config.configure(
            hover_color=T["bg_card_hover"], text_color=T["text_dim"])

        # Botão conectar — sempre verde no SGP, roxo no PV
        self.btn_conectar.configure(
            fg_color=T["accent"], hover_color=T["accent_hover"],
            text_color="#fffbe7" if tema == "sgp" else "#FFFFFF",
            font=ctk.CTkFont(family="Roboto", size=11, weight="bold"))
        self.status_lbl.configure(text_color=T["text_dim"],
            font=ctk.CTkFont(family="Roboto", size=9))

        # Tabs
        if tema == "sgp":
            self.btn_tab_sgp.configure(fg_color=T["tab_active"],
                hover_color=T["bg_card_hover"],
                border_color=T["accent"], text_color=T["text"],
                border_width=2)
            self.btn_tab_pv.configure(fg_color=T["bg_card"],
                hover_color=T["bg_card_hover"],
                border_color=T["border"], text_color=T["text_dim"],
                border_width=1)
        else:
            self.btn_tab_pv.configure(fg_color=T["accent_dark"],
                hover_color=T["accent_hover"],
                border_color=T["accent"], text_color=T["accent"],
                border_width=2)
            self.btn_tab_sgp.configure(fg_color=T["bg_card"],
                hover_color=T["bg_card_hover"],
                border_color=T["border"], text_color=T["text_dim"],
                border_width=1)
        self._animating = False

    # ══════════════════════════════════════════════════════════
    #  ABA SGP (cinza como o site SGP)
    # ══════════════════════════════════════════════════════════

    def _construir_aba_sgp(self):
        T = THEME_SGP
        tab = self.painel_sgp

        # Toolbar
        toolbar = ctk.CTkFrame(tab, fg_color="transparent", height=32)
        toolbar.pack(fill="x", padx=8, pady=(6, 4))

        # Botão Nova Bind = azul como "Cadastrar Ocorrência" do SGP
        ctk.CTkButton(toolbar, text="Cadastrar Bind", height=28, corner_radius=4,
            command=self._nova_bind,
            fg_color=T["blue"], hover_color=T["blue_hover"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=T["font"], size=10, weight="bold"),
        ).pack(side="left", padx=2)

        self.hotkey_lbl = ctk.CTkLabel(toolbar, text="",
            font=ctk.CTkFont(family=T["font"], size=9), text_color=T["text_dim"])
        self.hotkey_lbl.pack(side="right", padx=6)

        # Header da tabela (como SGP)
        hdr = ctk.CTkFrame(tab, fg_color=T["bg_card"], height=28, corner_radius=0)
        hdr.pack(fill="x", padx=6, pady=(0, 0))
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="Tecla", width=50, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=9, weight="bold"),
            text_color=T["text_dim"]).pack(side="left", padx=(10, 0))
        ctk.CTkLabel(hdr, text="Descrição", anchor="w",
            font=ctk.CTkFont(family=T["font"], size=9, weight="bold"),
            text_color=T["text_dim"]).pack(side="left", padx=(10, 0), fill="x", expand=True)
        ctk.CTkLabel(hdr, text="Tipo", width=65, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=9, weight="bold"),
            text_color=T["text_dim"]).pack(side="left")
        ctk.CTkLabel(hdr, text="", width=85).pack(side="right")

        # Lista de binds (estilo tabela de ocorrências)
        self.frame_binds = ctk.CTkScrollableFrame(tab, corner_radius=0,
            fg_color=T["bg"], scrollbar_button_color=T["border"],
            scrollbar_button_hover_color=T["bg_card_hover"])
        self.frame_binds.pack(fill="both", expand=True, padx=6, pady=0)
        self._atualizar_binds()

        ctk.CTkButton(tab, text="↺  Recarregar", height=24, corner_radius=4,
            command=self._recarregar, fg_color="transparent",
            hover_color=T["bg_card_hover"],
            font=ctk.CTkFont(family=T["font"], size=9), text_color=T["text_dim"]
        ).pack(fill="x", padx=8, pady=(2, 4))

    # ══════════════════════════════════════════════════════════
    #  ABA PAPERVINES (claro / roxo)
    # ══════════════════════════════════════════════════════════

    def _construir_aba_papervines(self):
        T = THEME_PV
        tab = self.painel_pv
        pv_cfg = self.cfg.get_papervines()

        # ── Scrollable para caber tudo ──
        scroll = ctk.CTkScrollableFrame(tab, fg_color=T["bg"], corner_radius=0,
            scrollbar_button_color=T["border"])
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # ════════════════════════════════════════
        #  SEÇÃO 1: Atendimento automático
        # ════════════════════════════════════════
        tf = ctk.CTkFrame(scroll, fg_color=T["accent_dark"], corner_radius=10)
        tf.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(tf, text="◈  Atendimento Automático",
            font=ctk.CTkFont(family=T["font"], size=12, weight="bold"),
            text_color=T["accent"]).pack(anchor="w", padx=14, pady=(8, 2))
        ctk.CTkLabel(tf, text="Novos → Iniciar → Saudação → Enviar",
            font=ctk.CTkFont(family=T["font"], size=9),
            text_color=T["text_dim"]).pack(anchor="w", padx=14, pady=(0, 8))

        # Card atendimento
        card1 = ctk.CTkFrame(scroll, fg_color=T["bg_card"], border_color=T["border"],
            border_width=1, corner_radius=10)
        card1.pack(fill="x", padx=8, pady=4)

        rk = ctk.CTkFrame(card1, fg_color="transparent")
        rk.pack(fill="x", padx=14, pady=(10, 4))
        ctk.CTkLabel(rk, text="Tecla:", width=48, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text_dim"]).pack(side="left")
        self.pv_key = ctk.CTkEntry(rk, width=55, corner_radius=6,
            font=ctk.CTkFont(family=T["font"], size=11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["accent"])
        self.pv_key.insert(0, pv_cfg.get("bind_key", "F9"))
        self.pv_key.pack(side="left", padx=(0, 14))
        ctk.CTkLabel(rk, text="Delay:", width=42, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text_dim"]).pack(side="left")
        self.pv_delay = ctk.CTkEntry(rk, width=50, corner_radius=6,
            font=ctk.CTkFont(family=T["font"], size=11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["accent"])
        self.pv_delay.insert(0, str(pv_cfg.get("delay_entre_clientes_ms", 1500)))
        self.pv_delay.pack(side="left")
        ctk.CTkLabel(rk, text="ms",
            font=ctk.CTkFont(family=T["font"], size=9),
            text_color=T["text_dim"]).pack(side="left", padx=3)

        ctk.CTkLabel(card1, text="Saudação:", anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text_dim"]).pack(fill="x", padx=14, pady=(4, 2))
        self.pv_saudacao = ctk.CTkTextbox(card1, height=55, corner_radius=8,
            font=ctk.CTkFont(family=T["font"], size=11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["text"])
        self.pv_saudacao.insert("1.0", pv_cfg.get("saudacao", ""))
        self.pv_saudacao.pack(fill="x", padx=14, pady=(0, 6))

        ctk.CTkButton(card1, text="Salvar", height=26, corner_radius=6,
            command=self._salvar_papervines,
            fg_color=T["bg_input"], hover_color=T["bg_card_hover"],
            text_color=T["accent_dim"],
            font=ctk.CTkFont(family=T["font"], size=10),
            border_width=1, border_color=T["border"]
        ).pack(fill="x", padx=14, pady=(0, 10))

        # Botão iniciar atendimento
        self.pv_btn = ctk.CTkButton(scroll,
            text="▶  Iniciar Atendimento", height=36,
            command=self._toggle_papervines, corner_radius=8,
            fg_color=T["accent"], hover_color=T["accent_dim"],
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=T["font"], size=11, weight="bold"))
        self.pv_btn.pack(fill="x", padx=8, pady=(4, 2))

        # ════════════════════════════════════════
        #  SEÇÃO 2: Transferência em massa
        # ════════════════════════════════════════
        tf2 = ctk.CTkFrame(scroll, fg_color=T["accent_dark"], corner_radius=10)
        tf2.pack(fill="x", padx=8, pady=(10, 4))
        ctk.CTkLabel(tf2, text="→  Transferir Clientes",
            font=ctk.CTkFont(family=T["font"], size=12, weight="bold"),
            text_color=T["accent"]).pack(anchor="w", padx=14, pady=(8, 2))
        ctk.CTkLabel(tf2, text="Novos → Transferir → Departamento → Repetir",
            font=ctk.CTkFont(family=T["font"], size=9),
            text_color=T["text_dim"]).pack(anchor="w", padx=14, pady=(0, 8))

        # Card transferência
        card2 = ctk.CTkFrame(scroll, fg_color=T["bg_card"], border_color=T["border"],
            border_width=1, corner_radius=10)
        card2.pack(fill="x", padx=8, pady=4)

        rk2 = ctk.CTkFrame(card2, fg_color="transparent")
        rk2.pack(fill="x", padx=14, pady=(10, 4))
        ctk.CTkLabel(rk2, text="Tecla:", width=48, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text_dim"]).pack(side="left")
        self.pv_transfer_key = ctk.CTkEntry(rk2, width=55, corner_radius=6,
            font=ctk.CTkFont(family=T["font"], size=11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["accent"])
        self.pv_transfer_key.insert(0, pv_cfg.get("transfer_bind_key", "F10"))
        self.pv_transfer_key.pack(side="left", padx=(0, 14))

        ctk.CTkLabel(card2, text="Departamento destino:", anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text_dim"]).pack(fill="x", padx=14, pady=(4, 2))
        self.pv_transfer_dept = ctk.CTkEntry(card2, height=32, corner_radius=6,
            font=ctk.CTkFont(family=T["font"], size=11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["text"],
            placeholder_text="Ex: Entrada Central, Cancelamento...")
        self.pv_transfer_dept.insert(0, pv_cfg.get("transfer_dept", "Entrada Central"))
        self.pv_transfer_dept.pack(fill="x", padx=14, pady=(0, 6))

        ctk.CTkButton(card2, text="Salvar", height=26, corner_radius=6,
            command=self._salvar_transfer,
            fg_color=T["bg_input"], hover_color=T["bg_card_hover"],
            text_color=T["accent_dim"],
            font=ctk.CTkFont(family=T["font"], size=10),
            border_width=1, border_color=T["border"]
        ).pack(fill="x", padx=14, pady=(0, 10))

        # Botão transferir
        self.pv_transfer_btn = ctk.CTkButton(scroll,
            text="→  Transferir Clientes", height=36,
            command=self._toggle_transfer, corner_radius=8,
            fg_color="#EF4444", hover_color="#DC2626",
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=T["font"], size=11, weight="bold"))
        self.pv_transfer_btn.pack(fill="x", padx=8, pady=(4, 2))

        # ════════════════════════════════════════
        #  STATUS + LOG (compartilhado)
        # ════════════════════════════════════════
        self.pv_status = ctk.CTkLabel(scroll, text="○  Parado",
            font=ctk.CTkFont(family=T["font"], size=10), text_color=T["text_dim"])
        self.pv_status.pack(anchor="w", padx=14, pady=(4, 2))

        self.pv_log = ctk.CTkTextbox(scroll, height=80, corner_radius=8,
            font=ctk.CTkFont(family="Consolas", size=9),
            fg_color=T["bg_card"], border_color=T["border"],
            text_color=T["text_dim"], state="disabled", border_width=1)
        self.pv_log.pack(fill="x", padx=8, pady=(2, 6))

        self.papervines._status_callback = self._pv_atualizar_status

    # ══════════════════════════════════════════════════════════
    #  PAPERVINES CONTROLS
    # ══════════════════════════════════════════════════════════

    def _salvar_papervines(self):
        try:
            self.cfg.update_papervines(
                bind_key=self.pv_key.get().strip(),
                delay_entre_clientes_ms=int(self.pv_delay.get().strip()),
                saudacao=self.pv_saudacao.get("1.0", "end-1c").strip(),
            )
            self._registrar_hotkeys()
            messagebox.showinfo("Salvo", "Configurações salvas!")
        except ValueError:
            messagebox.showerror("Erro", "Delay precisa ser número inteiro.")

    def _toggle_papervines(self):
        T = THEME_PV
        if self.papervines.esta_rodando():
            self.papervines.parar()
            self.pv_btn.configure(text="▶  Iniciar atendimento",
                fg_color=T["accent"], hover_color=T["accent_dim"])
            self.pv_status.configure(text="○  Parando...", text_color="#EF9F27")
        else:
            self._iniciar_papervines()

    def _iniciar_papervines(self):
        if self.sgp.driver is None or not self.sgp.esta_conectado():
            messagebox.showwarning("Chrome", "Conecte ao Chrome primeiro.")
            return
        saudacao = self.pv_saudacao.get("1.0", "end-1c").strip()
        delay = int(self.pv_delay.get().strip() or "1500")
        if not saudacao:
            messagebox.showwarning("Atenção", "Digite uma saudação.")
            return
        T = THEME_PV
        self.pv_btn.configure(text="◼  Parar atendimento",
            fg_color="#DC2626", hover_color="#B91C1C")
        self.pv_status.configure(text="●  Rodando...", text_color=T["accent"])
        def run():
            self.papervines.executar_loop(saudacao, delay)
            self.after(0, lambda: self.pv_btn.configure(
                text="▶  Iniciar atendimento",
                fg_color=T["accent"], hover_color=T["accent_dim"]))
            self.after(0, lambda: self.pv_status.configure(
                text="○  Parado", text_color=T["text_dim"]))
        threading.Thread(target=run, daemon=True).start()

    def _pv_atualizar_status(self, msg):
        def update():
            try:
                self.pv_status.configure(text=f"●  {msg}")
                self.pv_log.configure(state="normal")
                ts = datetime.now().strftime("%H:%M:%S")
                self.pv_log.insert("end", f"[{ts}] {msg}\n")
                self.pv_log.see("end")
                self.pv_log.configure(state="disabled")
            except Exception:
                pass
        self.after(0, update)

    def _salvar_transfer(self):
        try:
            self.cfg.update_papervines(
                transfer_dept=self.pv_transfer_dept.get().strip(),
                transfer_bind_key=self.pv_transfer_key.get().strip(),
            )
            self._registrar_hotkeys()
            messagebox.showinfo("Salvo", "Configurações de transferência salvas!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _toggle_transfer(self):
        T = THEME_PV
        if self.papervines.esta_rodando():
            self.papervines.parar()
            self.pv_transfer_btn.configure(text="→  Transferir Clientes",
                fg_color="#EF4444", hover_color="#DC2626")
            self.pv_btn.configure(state="normal")
            self.pv_status.configure(text="○  Parando...", text_color="#EF9F27")
        else:
            self._iniciar_transfer()

    def _iniciar_transfer(self):
        if self.sgp.driver is None or not self.sgp.esta_conectado():
            messagebox.showwarning("Chrome", "Conecte ao Chrome primeiro.")
            return
        dept = self.pv_transfer_dept.get().strip()
        if not dept:
            messagebox.showwarning("Atenção", "Digite o departamento destino.")
            return
        delay = int(self.pv_delay.get().strip() or "1500")

        T = THEME_PV
        self.pv_transfer_btn.configure(text="◼  Parar Transferência",
            fg_color="#991B1B", hover_color="#7F1D1D")
        self.pv_btn.configure(state="disabled")
        self.pv_status.configure(text=f"●  Transferindo → {dept}", text_color="#EF4444")

        def run():
            self.papervines.executar_transferencia(dept, delay)
            self.after(0, lambda: self.pv_transfer_btn.configure(
                text="→  Transferir Clientes",
                fg_color="#EF4444", hover_color="#DC2626"))
            self.after(0, lambda: self.pv_btn.configure(state="normal"))
            self.after(0, lambda: self.pv_status.configure(
                text="○  Parado", text_color=T["text_dim"]))

        threading.Thread(target=run, daemon=True).start()

    # ══════════════════════════════════════════════════════════
    #  SGP BIND CARDS
    # ══════════════════════════════════════════════════════════

    def _atualizar_binds(self):
        for w in self.frame_binds.winfo_children():
            w.destroy()
        for bind in self.cfg.get_binds():
            self._card_bind(bind)

    def _card_bind(self, bind: dict):
        T = THEME_SGP
        ativo = bind.get("enabled", True)
        tipos = bind.get("types", ["text"])
        nomes = {"text":"Texto","text_ocr":"Texto+Nº",
                 "sgp_ocorrencia":"SGP","link_pagamento":"Link"}
        tipo_str = " + ".join(nomes.get(t, t) for t in tipos)

        # Status color: verde=Aberta, vermelho=Encerrada (como na tabela SGP)
        status_color = T["accent"] if ativo else T["red"]
        status_text  = "Ativo" if ativo else "Inativo"

        # Row (estilo tabela de ocorrências do SGP)
        row = ctk.CTkFrame(self.frame_binds, corner_radius=0, height=42,
            fg_color=T["bg_card"] if ativo else T["bg"],
            border_width=0)
        row.pack(fill="x", pady=0, padx=0)
        row.pack_propagate(False)

        # Borda inferior (como as linhas da tabela SGP)
        ctk.CTkFrame(row, height=1, fg_color=T["border"],
            corner_radius=0).pack(side="bottom", fill="x")

        # Indicador de status lateral
        ctk.CTkFrame(row, width=3, corner_radius=0,
            fg_color=status_color).pack(side="left", fill="y")

        # Tecla
        ctk.CTkLabel(row, text=bind['key'], width=45, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10, weight="bold"),
            text_color=T["text"] if ativo else T["text_dim"]
        ).pack(side="left", padx=(10, 0))

        # Nome
        ctk.CTkLabel(row, text=bind['name'], anchor="w",
            font=ctk.CTkFont(family=T["font"], size=10),
            text_color=T["text"] if ativo else T["text_dim"]
        ).pack(side="left", padx=(10, 0), fill="x", expand=True)

        # Tipo
        ctk.CTkLabel(row, text=tipo_str, width=65, anchor="w",
            font=ctk.CTkFont(family=T["font"], size=9),
            text_color=T["text_dim"]
        ).pack(side="left")

        # Switch
        var = ctk.BooleanVar(value=ativo)
        ctk.CTkSwitch(row, variable=var, text="", width=36, height=18,
            progress_color=T["accent"], button_color="#CCCCCC",
            fg_color=T["red_dim"] if not ativo else T["bg_input"],
            command=lambda b=bind, v=var: self._toggle(b, v)
        ).pack(side="right", padx=(4, 6))

        # Editar
        ctk.CTkButton(row, text="✎", width=24, height=24, corner_radius=4,
            command=lambda b=bind: self._editar_bind(b),
            fg_color="transparent", hover_color=T["bg_card_hover"],
            font=ctk.CTkFont(size=11), text_color=T["text_dim"]
        ).pack(side="right", padx=1)

        # Deletar
        ctk.CTkButton(row, text="✕", width=24, height=24, corner_radius=4,
            command=lambda b=bind: self._deletar_bind(b),
            fg_color="transparent", hover_color=T["red_dim"],
            font=ctk.CTkFont(size=10), text_color="#884444"
        ).pack(side="right", padx=0)

    # ══════════════════════════════════════════════════════════
    #  CHROME / AÇÕES
    # ══════════════════════════════════════════════════════════

    def _conectar_chrome(self):
        self.btn_conectar.configure(text="Conectando...", state="disabled")
        self.update()
        def tentar():
            ok = self.sgp.conectar()
            self.after(0, lambda: self._pos_conexao(ok))
        threading.Thread(target=tentar, daemon=True).start()

    def _pos_conexao(self, ok: bool):
        T = self.T
        self.btn_conectar.configure(state="normal")
        if ok:
            self.btn_conectar.configure(text="◆  Conectado",
                fg_color="#28a745", hover_color="#2fbd4f",
                text_color="#fffbe7")
            self.status_lbl.configure(text="●  Conectado", text_color="#28a745")
        else:
            self.btn_conectar.configure(text="◇  Conectar ao Chrome",
                fg_color="#28a745", hover_color="#2fbd4f",
                text_color="#fffbe7")
            self.status_lbl.configure(text="○  Desconectado", text_color=T["text_dim"])

    def _toggle(self, bind, var):
        self.cfg.update_bind(bind["id"], enabled=var.get())
        self._recarregar()

    def _editar_bind(self, bind):
        BindEditorWindow(self, bind, self.cfg, self._recarregar)

    def _deletar_bind(self, bind):
        if messagebox.askyesno("Deletar",
                f"Deletar '{bind['name']}' [{bind['key']}]?"):
            self.cfg.data["binds"] = [
                b for b in self.cfg.data["binds"] if b["id"] != bind["id"]]
            self.cfg.save()
            self._recarregar()

    def _nova_bind(self):
        novo = {
            "id": f"bind_{int(time.time())}",
            "key": "F9", "name": "Nova bind",
            "types": ["text"], "message": "",
            "enabled": True, "cor": "#888780",
            "sgp_tipo_filtro": "", "sgp_origem_filtro": "whatsapp",
            "sgp_desmarcar_os": True, "sgp_auto_cadastrar": False,
            "sgp_auto_cadastrar_os": False
        }
        self.cfg.data["binds"].append(novo)
        self.cfg.save()
        BindEditorWindow(self, novo, self.cfg, self._recarregar)

    def _abrir_config(self):
        ConfigWindow(self, self.cfg)

    def _recarregar(self):
        self.hotkeys.limpar()
        self.cfg.reload()
        self._atualizar_binds()
        self._registrar_hotkeys()

    # ── Hotkeys ──────────────────────────────────────────────

    def _registrar_hotkeys(self):
        if not KEYBOARD_OK:
            self.hotkey_lbl.configure(text="⚠ 'keyboard' ausente", text_color="#EF9F27")
            return
        self.hotkeys.limpar()
        ativos = 0
        for bind in self.cfg.get_binds():
            if not bind.get("enabled", True):
                continue
            ativos += 1
            def make(b):
                def cb():
                    threading.Thread(target=self.executor.executar, args=(b,), daemon=True).start()
                return cb
            self.hotkeys.registrar(bind["key"].lower(), make(bind))

        pv_cfg = self.cfg.get_papervines()
        if pv_cfg.get("enabled", True):
            # Atendimento
            pv_key = pv_cfg.get("bind_key", "F9")
            def pv_cb():
                self.after(0, self._toggle_papervines)
            self.hotkeys.registrar(pv_key.lower(), pv_cb)
            ativos += 1

            # Transferência
            tr_key = pv_cfg.get("transfer_bind_key", "F10")
            def tr_cb():
                self.after(0, self._toggle_transfer)
            self.hotkeys.registrar(tr_key.lower(), tr_cb)
            ativos += 1

        self.hotkey_lbl.configure(text=f"● {ativos} hotkeys",
            text_color=THEME_SGP["text_dim"])

    def _loop_topo(self):
        self.attributes("-topmost", True)
        self.after(2500, self._loop_topo)

    def _fechar(self):
        self.papervines.parar()
        self.hotkeys.limpar()
        self.destroy()

# ════════════════════════════════════════════════════════════
#  UI: EDITOR DE BIND
# ════════════════════════════════════════════════════════════

class BindEditorWindow(ctk.CTkToplevel):

    def __init__(self, parent, bind: dict, cfg: ConfigManager, on_save):
        super().__init__(parent)
        self.bind_data = bind.copy()
        self.cfg       = cfg
        self.on_save   = on_save
        T = THEME

        self.title(f"Editar — {bind['name']}")
        self.geometry("480x700")
        self.attributes("-topmost", True)
        self.grab_set()
        self.resizable(False, True)
        self.configure(fg_color=T["bg"])
        self._build()

    def _build(self):
        T = THEME
        pad = {"padx": 16, "pady": 4}
        tipos_ativos = self.bind_data.get("types", ["text"])

        # ── Nome ──
        ctk.CTkLabel(self, text="Nome", anchor="w", font=_font(10),
                     text_color=T["text_dim"]).pack(fill="x", **pad)
        self.nome = ctk.CTkEntry(self, font=_font(11), fg_color=T["bg_input"],
                                 border_color=T["border"], text_color=T["text"])
        self.nome.insert(0, self.bind_data.get("name", ""))
        self.nome.pack(fill="x", **pad)

        # ── Tecla + Cor ──
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", **pad)
        ctk.CTkLabel(row, text="Tecla", width=55, anchor="w", font=_font(10),
                     text_color=T["text_dim"]).pack(side="left")
        self.key = ctk.CTkEntry(row, width=80, font=_font(11), fg_color=T["bg_input"],
                                border_color=T["border"], text_color=T["accent"])
        self.key.insert(0, self.bind_data.get("key", ""))
        self.key.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(row, text="Cor", width=35, anchor="w", font=_font(10),
                     text_color=T["text_dim"]).pack(side="left")
        self.cor = ctk.CTkEntry(row, width=90, font=_font(11), fg_color=T["bg_input"],
                                border_color=T["border"], text_color=T["text"])
        self.cor.insert(0, self.bind_data.get("cor", T["accent"]))
        self.cor.pack(side="left")

        # ── Tipos de ação ──
        ctk.CTkLabel(
            self, text="Tipos de ação", anchor="w",
            font=_font(11, True), text_color=T["accent"]
        ).pack(fill="x", padx=16, pady=(10, 2))

        self.tipo_vars = {}
        frame_tipos = ctk.CTkFrame(self, fg_color="transparent")
        frame_tipos.pack(fill="x", padx=16)
        for val, label in TIPOS_ACAO:
            var = ctk.BooleanVar(value=(val in tipos_ativos))
            self.tipo_vars[val] = var
            ctk.CTkCheckBox(
                frame_tipos, text=label, variable=var,
                font=_font(11), text_color=T["text"],
                fg_color=T["bg_input"], border_color=T["border"],
                checkmark_color=T["accent"], hover_color=T["accent_dark"]
            ).pack(anchor="w", pady=2)

        # ── Mensagem ──
        ctk.CTkLabel(
            self, text="Mensagem  —  {numero}=tel SMS  {mes}  {link}",
            anchor="w", font=_font(9), text_color=T["text_dim"]
        ).pack(fill="x", padx=16, pady=(10, 2))

        self.msg = ctk.CTkTextbox(
            self, height=80, font=_font(11),
            fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["text"]
        )
        self.msg.insert("1.0", self.bind_data.get("message", ""))
        self.msg.pack(fill="x", padx=16, pady=2)

        # ── Filtros SGP ──
        ctk.CTkLabel(
            self, text="Filtros SGP", anchor="w",
            font=_font(11, True), text_color=T["accent"]
        ).pack(fill="x", padx=16, pady=(10, 2))

        row2 = ctk.CTkFrame(self, fg_color="transparent")
        row2.pack(fill="x", padx=16)
        ctk.CTkLabel(row2, text="Tipo:", width=48, anchor="w", font=_font(10),
                     text_color=T["text_dim"]).pack(side="left")
        self.sgp_tipo = ctk.CTkEntry(
            row2, width=150, placeholder_text="sus / finan / reparo",
            font=_font(11), fg_color=T["bg_input"], border_color=T["border"],
            text_color=T["text"]
        )
        self.sgp_tipo.insert(0, self.bind_data.get("sgp_tipo_filtro", ""))
        self.sgp_tipo.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(row2, text="Origem:", width=52, anchor="w", font=_font(10),
                     text_color=T["text_dim"]).pack(side="left")
        self.sgp_origem = ctk.CTkEntry(
            row2, width=110, font=_font(11), fg_color=T["bg_input"],
            border_color=T["border"], text_color=T["text"]
        )
        self.sgp_origem.insert(0, self.bind_data.get("sgp_origem_filtro", "whatsapp"))
        self.sgp_origem.pack(side="left")

        ctk.CTkLabel(
            self, text="VMA (supersonic) troca 'sus' → 'finan' automaticamente",
            anchor="w", font=_font(9), text_color="#444"
        ).pack(fill="x", padx=16, pady=(2, 0))

        # ── Opções SGP (checkboxes) ──
        ctk.CTkLabel(
            self, text="Opções do formulário", anchor="w",
            font=_font(11, True), text_color=T["accent"]
        ).pack(fill="x", padx=16, pady=(10, 2))

        frame_opts = ctk.CTkFrame(self, fg_color="transparent")
        frame_opts.pack(fill="x", padx=16)

        self.desmarcar_os_var = ctk.BooleanVar(
            value=self.bind_data.get("sgp_desmarcar_os", True)
        )
        ctk.CTkCheckBox(
            frame_opts, text="Desmarcar 'Gerar OS'",
            variable=self.desmarcar_os_var, font=_font(11),
            text_color=T["text"], fg_color=T["bg_input"],
            border_color=T["border"], checkmark_color=T["accent"],
            hover_color=T["accent_dark"]
        ).pack(anchor="w", pady=2)

        self.auto_cadastrar_var = ctk.BooleanVar(
            value=self.bind_data.get("sgp_auto_cadastrar", False)
        )
        ctk.CTkCheckBox(
            frame_opts, text="Cadastrar automaticamente",
            variable=self.auto_cadastrar_var, font=_font(11),
            text_color=T["text"], fg_color=T["bg_input"],
            border_color=T["border"], checkmark_color=T["accent"],
            hover_color=T["accent_dark"]
        ).pack(anchor="w", pady=2)

        self.auto_cadastrar_os_var = ctk.BooleanVar(
            value=self.bind_data.get("sgp_auto_cadastrar_os", False)
        )
        ctk.CTkCheckBox(
            frame_opts, text="Cadastrar + OS (motivo: corretiva)",
            variable=self.auto_cadastrar_os_var, font=_font(11),
            text_color=T["text"], fg_color=T["bg_input"],
            border_color=T["border"], checkmark_color=T["accent"],
            hover_color=T["accent_dark"]
        ).pack(anchor="w", pady=2)
        ctk.CTkLabel(
            frame_opts, text="Cadastra ocorrência → preenche OS corretiva → cadastra OS",
            anchor="w", font=_font(9), text_color="#444"
        ).pack(anchor="w", padx=20, pady=(0, 2))

        # ── Botão Salvar ──
        ctk.CTkButton(
            self, text="Salvar", height=40, command=self._salvar,
            fg_color=T["accent_dark"], hover_color=T["accent_hover"],
            text_color=T["accent"], font=_font(12, True),
            border_width=1, border_color=T["accent"]
        ).pack(fill="x", padx=16, pady=(14, 10))

    def _salvar(self):
        tipos_sel = [val for val, var in self.tipo_vars.items() if var.get()]
        if not tipos_sel:
            messagebox.showwarning("Atenção", "Selecione ao menos um tipo de ação.")
            return

        self.cfg.update_bind(
            self.bind_data["id"],
            name              = self.nome.get(),
            key               = self.key.get().strip(),
            cor               = self.cor.get().strip(),
            types             = tipos_sel,
            message           = self.msg.get("1.0", "end-1c"),
            sgp_tipo_filtro   = self.sgp_tipo.get(),
            sgp_origem_filtro = self.sgp_origem.get(),
            sgp_desmarcar_os  = self.desmarcar_os_var.get(),
            sgp_auto_cadastrar = self.auto_cadastrar_var.get(),
            sgp_auto_cadastrar_os = self.auto_cadastrar_os_var.get(),
        )
        self.on_save()
        self.destroy()

# ════════════════════════════════════════════════════════════
#  UI: CONFIGURAÇÕES
# ════════════════════════════════════════════════════════════

class ConfigWindow(ctk.CTkToplevel):

    def __init__(self, parent, cfg: ConfigManager):
        super().__init__(parent)
        self.cfg = cfg
        T = THEME
        self.title("Configurações")
        self.geometry("420x310")
        self.attributes("-topmost", True)
        self.grab_set()
        self.resizable(False, True)
        self.configure(fg_color=T["bg"])
        self._build()

    def _build(self):
        T = THEME
        ctk.CTkLabel(
            self, text="Configurações",
            font=_font(14, True), text_color=T["accent"]
        ).pack(pady=(16, 10))

        frame = ctk.CTkFrame(self, fg_color=T["bg_card"], border_color=T["border"],
                             border_width=1, corner_radius=8)
        frame.pack(fill="x", padx=16, pady=4)

        ctk.CTkLabel(frame, text="Porta de depuração do Chrome", anchor="w",
                     font=_font(10), text_color=T["text"]).pack(fill="x", padx=12, pady=(10, 2))
        self.porta = ctk.CTkEntry(frame, width=100, font=_font(11),
                                  fg_color=T["bg_input"], border_color=T["border"],
                                  text_color=T["accent"])
        self.porta.insert(0, str(self.cfg.get_sgp().get("debug_port", 9222)))
        self.porta.pack(anchor="w", padx=12)
        ctk.CTkLabel(frame, text="Padrão 9222", font=_font(9),
                     text_color=T["text_dim"], anchor="w").pack(fill="x", padx=12, pady=(2, 8))

        ctk.CTkLabel(frame, text="Delay entre ações (ms)", anchor="w",
                     font=_font(10), text_color=T["text"]).pack(fill="x", padx=12, pady=(4, 2))
        self.delay = ctk.CTkEntry(frame, width=100, font=_font(11),
                                  fg_color=T["bg_input"], border_color=T["border"],
                                  text_color=T["accent"])
        self.delay.insert(0, str(self.cfg.get_sgp().get("delay_ms", 150)))
        self.delay.pack(anchor="w", padx=12)
        ctk.CTkLabel(frame, text="Menor = mais rápido (padrão 150)",
                     font=_font(9), text_color=T["text_dim"], anchor="w"
        ).pack(fill="x", padx=12, pady=(2, 10))

        ctk.CTkButton(
            self, text="Salvar", height=38, command=self._salvar,
            fg_color=T["accent_dark"], hover_color=T["accent_hover"],
            text_color=T["accent"], font=_font(12, True),
            border_width=1, border_color=T["accent"]
        ).pack(fill="x", padx=16, pady=10)

    def _salvar(self):
        try:
            self.cfg.data["sgp"]["debug_port"] = int(self.porta.get())
            self.cfg.data["sgp"]["delay_ms"]   = int(self.delay.get())
            self.cfg.save()
            messagebox.showinfo("Salvo", "Configurações salvas!")
            self.destroy()
        except ValueError:
            messagebox.showerror("Erro", "Porta e delay precisam ser números inteiros.")

# ════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════

if __name__ == "__main__":
    faltando = []
    if not KEYBOARD_OK:  faltando.append("keyboard")
    if not PYAUTOGUI_OK: faltando.append("pyautogui / pyperclip")
    if not SELENIUM_OK:  faltando.append("selenium")
    if faltando:
        print("Bibliotecas faltando:", ", ".join(faltando))
        print("Execute instalar.bat")

    app = FloatingApp()
    app.mainloop()
