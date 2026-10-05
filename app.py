#!/usr/bin/env python3
"""
app.py - Servidor Web para el Portal de Pedidos de Clientes y Panel de Cocina.
Sincronización automática en tiempo real con el archivo Excel en el Escritorio.
"""

import os
import re
import time
from datetime import datetime
import json
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, jsonify, send_file
from viandas_engine import (
    Catalogo, GestorPedidos, GeneradorExcel, GestorUsuarios,
    EXCEL_OUTPUT, DIAS_ORDEN, normalize_text
)

app = Flask(__name__)
UPLOAD_FOLDER = Path(__file__).resolve().parent / "static" / "uploads"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp', 'gif'}
CONFIG_FILE = Path(__file__).resolve().parent / "config.json"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                if "alias_mercadopago" not in cfg:
                    cfg["alias_mercadopago"] = "julia.rabovich"
                if "link_pago_mp" not in cfg:
                    cfg["link_pago_mp"] = ""
                return cfg
        except Exception:
            pass
    return {
        "modo_pedidos": "abierto",
        "limite_dia": "miercoles",
        "limite_hora": 12,
        "whatsapp_mama": "",
        "alias_mercadopago": "julia.rabovich",
        "link_pago_mp": ""
    }

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

def estado_horario_pedidos():
    """
    Determina si los pedidos están abiertos según la configuración y el horario.
    Regla: Abierto de Lunes a Miércoles hasta las 12 PM.
    Jueves a Domingo: Cerrado (la cocina ya está preparando).
    """
    cfg = get_config()
    modo = cfg.get("modo_pedidos", "abierto")
    if modo == "abierto":
        return True, "Abierto (Modo extendido / pruebas)", cfg
    if modo == "cerrado":
        return False, "Pedidos cerrados por la cocina.", cfg

    # Modo AUTO
    now = datetime.now()
    wd = now.weekday()  # 0=Lunes, 1=Martes, 2=Miércoles, 3=Jueves, 4=Viernes, 5=Sábado, 6=Domingo
    if wd in [0, 1]:
        return True, "Abierto (cierra el Miércoles a las 12 PM)", cfg
    if wd == 2:
        if now.hour < 12:
            minutos = (11 - now.hour) * 60 + (60 - now.minute)
            horas = minutos // 60
            mins = minutos % 60
            return True, f"Abierto (cierra hoy a las 12 PM - quedan {horas}h {mins}m)", cfg
        else:
            return False, "Pedidos cerrados hoy Miércoles a las 12 PM. Se habilita nuevamente el próximo lunes.", cfg

    return False, "Pedidos cerrados por esta semana. La cocina ya está preparando las viandas. Se habilita el lunes.", cfg

def puede_cancelar_pedido(pedido):
    """Verifica si el pedido puede ser cancelado (antes del miércoles a las 12 PM)."""
    if pedido.get("entregado"):
        return False, "El pedido ya fue entregado."
    abierto, msg, cfg = estado_horario_pedidos()
    if abierto:
        return True, "Podés modificar o cancelar tu pedido hasta el Miércoles a las 12 PM."
    return False, "El plazo para cancelar finalizó el Miércoles a las 12 PM. Para cambios urgentes, por favor comunicate directamente por WhatsApp."

catalogo = Catalogo()
gestor = GestorPedidos(catalogo)
gestor_usuarios = GestorUsuarios()

@app.route("/")
def portal_cliente():
    """Portal de pedidos para clientes desde el celular con soporte de cookies y control de horario."""
    catalogo.load()
    gestor_usuarios.load()

    # Verificar horario
    pedidos_abiertos, mensaje_horario, cfg = estado_horario_pedidos()

    # Verificar si el cliente tiene una cookie guardada de su usuario
    user_id = request.cookies.get("viandas_user_id")
    usuario_actual = None
    if user_id:
        usuario_actual = gestor_usuarios.get_by_id(user_id)

    return render_template(
        "cliente.html",
        menu=catalogo.items,
        usuario_actual=usuario_actual,
        pedidos_abiertos=pedidos_abiertos,
        mensaje_horario=mensaje_horario,
        whatsapp_mama=cfg.get("whatsapp_mama", ""),
        alias_mercadopago=cfg.get("alias_mercadopago", "julia.rabovich"),
        link_pago_mp=cfg.get("link_pago_mp", ""),
        modo_pedidos=cfg.get("modo_pedidos", "abierto")
    )

@app.route("/cocina")
def panel_cocina():
    """Panel de control de cocina con resumen por día y cobros."""
    catalogo.load()
    gestor.load()
    pedidos_abiertos, estado_pedidos_msg, cfg = estado_horario_pedidos()

    # Agrupar pedidos para la cocina por día y plato
    dias_agrupados = {}
    total_viandas = 0
    total_dinero = 0.0
    total_pagado = 0.0

    for p in gestor.pedidos:
        d = p.get("dia", "Varios")
        plato = p.get("plato", "Sin nombre")
        cant = p.get("cantidad", 1)
        tot = p.get("total", 0.0)
        cli = p.get("cliente", "")
        pag = p.get("pagado", False)

        total_viandas += cant
        total_dinero += tot
        if pag:
            total_pagado += tot

        dias_agrupados.setdefault(d, {}).setdefault(plato, {"cant": 0, "clientes": []})
        dias_agrupados[d][plato]["cant"] += cant
        dias_agrupados[d][plato]["clientes"].append(f"{cli} (x{cant})" if cant > 1 else cli)

    def sort_dia_key(d):
        norm = normalize_text(d)
        return DIAS_ORDEN.get(norm, 99)

    cocina_ordenada = {}
    for d in sorted(dias_agrupados.keys(), key=sort_dia_key):
        cocina_ordenada[d] = dias_agrupados[d]

    total_pendiente = total_dinero - total_pagado

    return render_template(
        "cocina.html",
        cocina_por_dia=cocina_ordenada,
        pedidos=list(reversed(gestor.pedidos)),
        catalogo=catalogo.items,
        total_viandas=total_viandas,
        total_dinero=total_dinero,
        total_pagado=total_pagado,
        total_pendiente=total_pendiente,
        usuarios=gestor_usuarios.usuarios,
        pedidos_abiertos=pedidos_abiertos,
        estado_pedidos_msg=estado_pedidos_msg,
        config=cfg
    )

@app.route("/api/config", methods=["POST"])
def api_guardar_config():
    """Permite cambiar el modo de pedidos (Auto, Abierto, Cerrado) o WhatsApp de contacto."""
    data = request.get_json() or {}
    cfg = get_config()
    if "modo_pedidos" in data:
        cfg["modo_pedidos"] = data["modo_pedidos"]
    if "whatsapp_mama" in data:
        cfg["whatsapp_mama"] = data["whatsapp_mama"].strip()
    if "alias_mercadopago" in data:
        cfg["alias_mercadopago"] = data["alias_mercadopago"].strip()
    if "link_pago_mp" in data:
        cfg["link_pago_mp"] = data["link_pago_mp"].strip()
    save_config(cfg)
    abierto, msg, _ = estado_horario_pedidos()
    return jsonify({"ok": True, "config": cfg, "pedidos_abiertos": abierto, "mensaje": msg})

@app.route("/api/pedir", methods=["POST"])
def api_pedir():
    """Recibe los pedidos seleccionados por el cliente desde la web si está en horario."""
    abierto, msg_horario, cfg = estado_horario_pedidos()
    if not abierto:
        return jsonify({
            "ok": False,
            "error": msg_horario,
            "cerrado": True,
            "whatsapp_mama": cfg.get("whatsapp_mama", ""),
            "alias_mercadopago": cfg.get("alias_mercadopago", "julia.rabovich"),
            "link_pago_mp": cfg.get("link_pago_mp", "")
        }), 403

    data = request.get_json() or {}
    cliente = data.get("cliente", "").strip()
    whatsapp = data.get("whatsapp", "").strip()
    notas = data.get("notas", "").strip()
    items = data.get("items", [])

    if not cliente:
        return jsonify({"ok": False, "error": "Falta el nombre del cliente"}), 400

    if not items:
        return jsonify({"ok": False, "error": "No se enviaron viandas"}), 400

    recordar_usuario = bool(data.get("recordar_usuario", True))
    usuario = None
    if recordar_usuario and cliente and whatsapp:
        usuario = gestor_usuarios.registrar_o_actualizar(
            nombre=cliente,
            whatsapp=whatsapp,
            notas=notas
        )

    usuario_id = usuario["id"] if usuario else None

    pedidos_creados = []
    for it in items:
        dia = it.get("dia", "Lunes")
        plato = it.get("plato_nombre") or str(it.get("plato_id"))
        cantidad = int(it.get("cantidad", 1))

        ped = gestor.add_pedido(
            cliente=cliente,
            plato_query=plato,
            dia=dia,
            cantidad=cantidad,
            pagado=False,
            notas=notas,
            whatsapp=whatsapp,
            entregado=False,
            usuario_id=usuario_id
        )
        pedidos_creados.append(ped)

    # Actualizar Excel inmediatamente
    GeneradorExcel.export(catalogo, gestor)

    total_monto = sum(p.get("total", 0.0) for p in pedidos_creados)
    alias_mp = cfg.get("alias_mercadopago", "julia.rabovich")
    link_pago_mp = cfg.get("link_pago_mp", "")

    resp = jsonify({
        "ok": True,
        "registrados": len(pedidos_creados),
        "cliente": cliente,
        "usuario": usuario,
        "total_monto": total_monto,
        "alias_mercadopago": alias_mp,
        "link_pago_mp": link_pago_mp,
        "pedidos_ids": [p["id"] for p in pedidos_creados]
    })

    if usuario:
        # Guardar cookie por 365 días
        resp.set_cookie(
            "viandas_user_id",
            str(usuario["id"]),
            max_age=365 * 24 * 60 * 60,
            path="/",
            samesite="Lax"
        )

    return resp

@app.route("/api/menu/guardar", methods=["POST"])
def api_guardar_menu():
    """Agrega o edita un plato del catálogo, con soporte para subir fotos."""
    if request.is_json:
        data = request.get_json() or {}
        item_id = data.get("id")
        nombre = data.get("nombre", "").strip()
        precio = data.get("precio", 0)
        aliases_raw = data.get("aliases", "")
        foto_final = data.get("foto", "")
        dia_raw = data.get("dia") or data.get("dias") or "Todos"
        sabores_raw = data.get("sabores", "")
        descripcion = data.get("descripcion", "").strip()
        tipo = data.get("tipo", "").strip()
    else:
        item_id = request.form.get("id")
        nombre = request.form.get("nombre", "").strip()
        precio = request.form.get("precio", 0)
        aliases_raw = request.form.get("aliases", "")
        foto_final = request.form.get("foto_url", "")
        dia_raw = request.form.get("dia") or request.form.getlist("dias") or "Todos"
        sabores_raw = request.form.get("sabores", "")
        descripcion = request.form.get("descripcion", "").strip()
        tipo = request.form.get("tipo", "").strip()

        # Archivo subido
        if 'foto_archivo' in request.files:
            file = request.files['foto_archivo']
            if file and file.filename and allowed_file(file.filename):
                safe_name = secure_filename(file.filename)
                unique_name = f"{int(time.time())}_{safe_name}"
                dest = UPLOAD_FOLDER / unique_name
                file.save(dest)
                foto_final = f"/static/uploads/{unique_name}"

    if not nombre or not precio:
        return jsonify({"ok": False, "error": "El nombre y el precio son requeridos"}), 400

    aliases = [a.strip() for a in aliases_raw.split(",") if a.strip()] if isinstance(aliases_raw, str) else (aliases_raw or [])

    # Procesar días
    if isinstance(dia_raw, list):
        dias = [d.strip().capitalize() for d in dia_raw if d.strip()]
    elif isinstance(dia_raw, str):
        dias = [d.strip() for d in dia_raw.split(",") if d.strip()]
    else:
        dias = ["Todos"]
    if not dias:
        dias = ["Todos"]

    # Procesar sabores
    if isinstance(sabores_raw, list):
        sabores = [s.strip() for s in sabores_raw if s.strip()]
    elif isinstance(sabores_raw, str):
        sabores = [s.strip() for s in sabores_raw.split(",") if s.strip()]
    else:
        sabores = []

    catalogo.add_or_update_item(
        nombre=nombre,
        precio=float(precio),
        aliases=aliases,
        foto=foto_final,
        item_id=int(item_id) if item_id else None,
        dias=dias,
        sabores=sabores,
        descripcion=descripcion,
        tipo=tipo
    )

    GeneradorExcel.export(catalogo, gestor)
    return jsonify({"ok": True, "catalogo": catalogo.items})

@app.route("/api/menu/eliminar", methods=["POST"])
def api_eliminar_menu():
    """Elimina un plato del menú."""
    data = request.get_json() or {}
    item_id = data.get("id")
    if not item_id:
        return jsonify({"ok": False, "error": "ID requerido"}), 400

    deleted = catalogo.delete_item(int(item_id))
    if deleted:
        GeneradorExcel.export(catalogo, gestor)
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Plato no encontrado"}), 404

@app.route("/api/mis-pedidos")
def api_mis_pedidos():
    """Devuelve los pedidos del cliente con validación de límite hasta el Miércoles a las 12 PM."""
    user_id = request.cookies.get("viandas_user_id") or request.args.get("usuario_id")
    whatsapp = request.args.get("whatsapp", "")
    pedido_id = request.args.get("pedido_id")

    if not user_id and not whatsapp and not pedido_id:
        return jsonify({"ok": True, "pedidos": []})

    pedidos = gestor.get_pedidos_cliente(usuario_id=user_id, whatsapp=whatsapp)
    if pedido_id:
        try:
            pid = int(pedido_id)
            if not any(p.get("id") == pid for p in pedidos):
                p_extra = next((p for p in gestor.pedidos if p.get("id") == pid), None)
                if p_extra:
                    pedidos.append(p_extra)
        except Exception:
            pass
    pedidos_con_tiempo = []
    cfg = get_config()

    for p in pedidos:
        puede_cancelar, motivo = puede_cancelar_pedido(p)
        p_copy = dict(p)
        p_copy["puede_cancelar"] = puede_cancelar
        p_copy["motivo_cancelacion"] = motivo
        p_copy["whatsapp_mama"] = cfg.get("whatsapp_mama", "")
        p_copy["alias_mercadopago"] = cfg.get("alias_mercadopago", "julia.rabovich")
        p_copy["link_pago_mp"] = cfg.get("link_pago_mp", "")
        pedidos_con_tiempo.append(p_copy)

    return jsonify({"ok": True, "pedidos": pedidos_con_tiempo})

@app.route("/api/pedido/cancelar", methods=["POST"])
def api_cancelar_pedido():
    """Cancela un pedido si el usuario está dentro del plazo permitido (hasta el Miércoles a las 12 PM)."""
    data = request.get_json() or {}
    pedido_id = data.get("pedido_id")

    if not pedido_id:
        return jsonify({"ok": False, "error": "ID de pedido requerido"}), 400

    target = None
    for p in gestor.pedidos:
        if p.get("id") == int(pedido_id):
            target = p
            break

    if not target:
        return jsonify({"ok": False, "error": "Pedido no encontrado"}), 404

    es_admin = bool(data.get("admin", False))
    puede_cancelar, motivo = puede_cancelar_pedido(target)
    cfg = get_config()

    if not es_admin and not puede_cancelar:
        return jsonify({
            "ok": False,
            "error": motivo,
            "expiro": True,
            "whatsapp_mama": cfg.get("whatsapp_mama", ""),
            "pedido": target
        }), 400

    gestor.delete_pedido(int(pedido_id))
    GeneradorExcel.export(catalogo, gestor)

    return jsonify({
        "ok": True,
        "mensaje": f"Pedido #{pedido_id} cancelado correctamente.",
        "pedido_cancelado": target
    })

@app.route("/api/usuario/olvidar", methods=["POST"])
def api_olvidar_usuario():
    """Borra la cookie de sesión del cliente para cambiar de usuario o teléfono."""
    resp = jsonify({"ok": True})
    resp.delete_cookie("viandas_user_id", path="/")
    return resp

@app.route("/api/pedido/toggle", methods=["POST"])
def api_toggle_pedido():
    """Alterna el estado de pago o entrega de un pedido."""
    data = request.get_json() or {}
    pedido_id = data.get("id")
    tipo = data.get("tipo")  # 'pago' o 'entrega'

    if not pedido_id or not tipo:
        return jsonify({"ok": False, "error": "Parámetros incompletos"}), 400

    if tipo == "pago":
        res = gestor.toggle_pago(int(pedido_id))
    elif tipo == "entrega":
        res = gestor.toggle_entrega(int(pedido_id))
    else:
        return jsonify({"ok": False, "error": "Tipo inválido"}), 400

    if res:
        GeneradorExcel.export(catalogo, gestor)
        total_dinero = sum(p.get("total", 0.0) for p in gestor.pedidos)
        total_pagado = sum(p.get("total", 0.0) for p in gestor.pedidos if p.get("pagado"))
        total_pendiente = total_dinero - total_pagado
        return jsonify({
            "ok": True,
            "pedido": res,
            "total_dinero": total_dinero,
            "total_pagado": total_pagado,
            "total_pendiente": total_pendiente
        })
    return jsonify({"ok": False, "error": "Pedido no encontrado"}), 404

@app.route("/descargar-excel")
def descargar_excel():
    """Permite descargar el archivo Excel actualizado con un toque."""
    # Asegurar que esté recién generado
    GeneradorExcel.export(catalogo, gestor)
    return send_file(
        EXCEL_OUTPUT,
        as_attachment=True,
        download_name="viandas_semanales.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    # Escuchar en todas las interfaces para permitir acceso desde celulares en la red local
    port = int(os.environ.get("PORT", 5000))
    print(f"🚀 Servidor de Viandas iniciado:")
    print(f"👉 Portal Clientes: http://localhost:{port}/")
    print(f"👉 Panel Cocina:    http://localhost:{port}/cocina")
    app.run(host="0.0.0.0", port=port, debug=False)
