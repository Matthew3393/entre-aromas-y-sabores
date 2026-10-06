/**
 * supabase-client.js
 * Cliente Supabase y funciones de lógica compartida para Entre Aromas y Sabores
 */

const SUPABASE_URL = 'https://ngptxegjqosodwleusgb.supabase.co';
const SUPABASE_KEY = 'sb_publishable_gX_hBbA1EvP-DygZ2Mx-ng_CNbi1hEf';

const headers = {
  'apikey': SUPABASE_KEY,
  'Authorization': `Bearer ${SUPABASE_KEY}`,
  'Content-Type': 'application/json',
  'Prefer': 'return=representation'
};

const Api = {
  // CONFIGURACIÓN
  async getConfig() {
    try {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/configuracion?id=eq.1&select=*`, { headers });
      const data = await res.json();
      if (data && data.length > 0) return data[0];
    } catch (e) {
      console.error('Error cargando config:', e);
    }
    return {
      modo_pedidos: 'abierto',
      limite_dia: 'miercoles',
      limite_hora: 12,
      whatsapp_mama: '2944335900',
      alias_mercadopago: 'julia.rabovich',
      link_pago_mp: ''
    };
  },

  async updateConfig(changes) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/configuracion?id=eq.1`, {
      method: 'PATCH',
      headers,
      body: JSON.stringify(changes)
    });
    return res.ok;
  },

  // CATÁLOGO
  async getCatalogo() {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/catalogo?select=*&order=id.asc`, { headers });
    return await res.json();
  },

  async saveCatalogoItem(item) {
    if (item.id) {
      const id = item.id;
      const cleanItem = { ...item };
      delete cleanItem.id;
      const res = await fetch(`${SUPABASE_URL}/rest/v1/catalogo?id=eq.${id}`, {
        method: 'PATCH',
        headers,
        body: JSON.stringify(cleanItem)
      });
      return await res.json();
    } else {
      const cleanItem = { ...item };
      delete cleanItem.id;
      const res = await fetch(`${SUPABASE_URL}/rest/v1/catalogo`, {
        method: 'POST',
        headers,
        body: JSON.stringify(cleanItem)
      });
      return await res.json();
    }
  },

  async deleteCatalogoItem(id) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/catalogo?id=eq.${id}`, {
      method: 'DELETE',
      headers
    });
    return res.ok;
  },

  // USUARIOS
  async getUsuarioByWhatsapp(whatsapp) {
    if (!whatsapp) return null;
    const cleanWa = whatsapp.trim();
    const encoded = encodeURIComponent(cleanWa);
    const res = await fetch(`${SUPABASE_URL}/rest/v1/usuarios?whatsapp=eq.${encoded}&select=*`, { headers });
    const data = await res.json();
    return data && data.length > 0 ? data[0] : null;
  },

  async getUsuarios() {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/usuarios?select=*&order=nombre.asc`, { headers });
    return await res.json();
  },

  async registrarOUsuario(nombre, whatsapp, notas = '') {
    const existing = await this.getUsuarioByWhatsapp(whatsapp);
    const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 16);
    if (existing) {
      const updates = {
        nombre: nombre.trim(),
        ultimo_acceso: nowStr,
        pedidos_count: (existing.pedidos_count || 1) + 1
      };
      if (notas && notas.trim()) {
        updates.notas_frecuentes = notas.trim();
      }
      await fetch(`${SUPABASE_URL}/rest/v1/usuarios?id=eq.${existing.id}`, {
        method: 'PATCH',
        headers,
        body: JSON.stringify(updates)
      });
      return { ...existing, ...updates };
    } else {
      const nuevo = {
        nombre: nombre.trim(),
        whatsapp: whatsapp.trim(),
        direccion: '',
        notas_frecuentes: notas.trim(),
        pedidos_count: 1,
        fecha_registro: nowStr,
        ultimo_acceso: nowStr
      };
      const res = await fetch(`${SUPABASE_URL}/rest/v1/usuarios`, {
        method: 'POST',
        headers,
        body: JSON.stringify(nuevo)
      });
      const data = await res.json();
      return data && data[0] ? data[0] : nuevo;
    }
  },

  // PEDIDOS
  async getPedidos() {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/pedidos?select=*&order=id.desc`, { headers });
    return await res.json();
  },

  async getPedidosCliente(whatsapp) {
    if (!whatsapp) return [];
    const encoded = encodeURIComponent(whatsapp.trim());
    const res = await fetch(`${SUPABASE_URL}/rest/v1/pedidos?whatsapp=eq.${encoded}&select=*&order=id.desc`, { headers });
    return await res.json();
  },

  async crearPedidos(listaPedidos) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/pedidos`, {
      method: 'POST',
      headers,
      body: JSON.stringify(listaPedidos)
    });
    return await res.json();
  },

  async togglePedido(id, campo, valor) {
    const body = {};
    body[campo] = valor;
    const res = await fetch(`${SUPABASE_URL}/rest/v1/pedidos?id=eq.${id}`, {
      method: 'PATCH',
      headers,
      body: JSON.stringify(body)
    });
    return res.ok;
  },

  async cancelarPedido(id) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/pedidos?id=eq.${id}`, {
      method: 'DELETE',
      headers
    });
    return res.ok;
  }
};

// Control de horario
function calcularEstadoHorario(cfg) {
  const modo = (cfg && cfg.modo_pedidos) ? cfg.modo_pedidos : 'abierto';
  if (modo === 'abierto') {
    return {
      abierto: true,
      mensaje: 'Abierto (Recepción de pedidos habilitada)'
    };
  }
  if (modo === 'cerrado') {
    return {
      abierto: false,
      mensaje: 'Pedidos cerrados por la cocina.'
    };
  }

  // Modo AUTO: Lunes a Jueves 12:00 PM
  const now = new Date();
  const day = now.getDay(); // 0 Dom, 1 Lun, 2 Mar, 3 Mie, 4 Jue, 5 Vie, 6 Sab
  const hour = now.getHours();
  const minute = now.getMinutes();

  if (day === 1 || day === 2 || day === 3) { // Lunes, Martes o Miércoles
    return {
      abierto: true,
      mensaje: 'Abierto (cierra el Jueves a las 12 PM)'
    };
  }
  if (day === 4) { // Jueves
    if (hour < 12) {
      const minutosRestantes = (11 - hour) * 60 + (60 - minute);
      const h = Math.floor(minutosRestantes / 60);
      const m = minutosRestantes % 60;
      return {
        abierto: true,
        mensaje: `Abierto (cierra hoy a las 12 PM - quedan ${h}h ${m}m)`
      };
    } else {
      return {
        abierto: false,
        mensaje: 'Pedidos cerrados hoy Jueves a las 12 PM. Se habilita el lunes.'
      };
    }
  }

  return {
    abierto: false,
    mensaje: 'Pedidos cerrados por esta semana. La cocina ya está preparando las viandas. Se habilita el lunes.'
  };
}

function puedeCancelarPedido(pedido, cfg) {
  if (pedido.pagado) {
    return { puede: false, motivo: 'El pedido ya fue registrado como abonado y confirmado por la cocina.' };
  }
  if (pedido.entregado) {
    return { puede: false, motivo: 'El pedido ya fue entregado.' };
  }
  const estado = calcularEstadoHorario(cfg);
  if (estado.abierto) {
    return { puede: true, motivo: 'Podés modificar o cancelar tu pedido hasta el Jueves a las 12 PM.' };
  }
  return {
    puede: false,
    motivo: 'El plazo para cancelar finalizó el Jueves a las 12 PM. Para cambios urgentes, por favor comunicate directamente por WhatsApp.'
  };
}

window.SupabaseApi = Api;
window.calcularEstadoHorario = calcularEstadoHorario;
window.puedeCancelarPedido = puedeCancelarPedido;
