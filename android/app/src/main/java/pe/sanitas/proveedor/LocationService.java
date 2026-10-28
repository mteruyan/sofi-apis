package pe.sanitas.proveedor;

import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.app.Service;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.content.pm.ServiceInfo;
import android.location.Location;
import android.location.LocationListener;
import android.location.LocationManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.HandlerThread;
import android.os.IBinder;
import android.os.PowerManager;

import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

/**
 * Servicio en primer plano: mientras el doctor está "disponible" envía su ubicación al sistema cada 15 s
 * (con la app minimizada o el celular bloqueado) y avisa con una notificación cuando llega una oferta.
 */
public class LocationService extends Service implements LocationListener {
    private static final String CANAL_SERVICIO = "ubicacion";
    private static final String CANAL_OFERTAS = "ofertas_v1";
    private static final int ID_SERVICIO = 1;
    private static final int ID_OFERTA = 2;
    private static final long INTERVALO_MS = 15000;
    private static final long POSICION_MAX_EDAD_MS = 60000;

    private HandlerThread hilo;
    private Handler handler;
    private LocationManager lm;
    private PowerManager.WakeLock wakeLock;
    private volatile Location ultima;
    private long ultimaOfertaAvisada = -1;
    private long ultimoServicioAvisado = -1;
    private boolean corriendo = false;

    private final Runnable ciclo = new Runnable() {
        @Override public void run() {
            if (!corriendo) return;
            try { tick(); } catch (Exception ignorada) { }
            if (corriendo) handler.postDelayed(this, INTERVALO_MS);
        }
    };

    @Override public IBinder onBind(Intent intent) { return null; }

    @Override
    public int onStartCommand(Intent intent, int flags, int startId) {
        crearCanales();
        Notification n = notificacionServicio("Compartiendo tu ubicación para recibir solicitudes");
        if (Build.VERSION.SDK_INT >= 29) startForeground(ID_SERVICIO, n, ServiceInfo.FOREGROUND_SERVICE_TYPE_LOCATION);
        else startForeground(ID_SERVICIO, n);

        if (leerToken() == null) { stopSelf(); return START_NOT_STICKY; }

        if (!corriendo) {
            corriendo = true;
            hilo = new HandlerThread("sanitas-ubicacion");
            hilo.start();
            handler = new Handler(hilo.getLooper());
            PowerManager pm = (PowerManager) getSystemService(Context.POWER_SERVICE);
            wakeLock = pm.newWakeLock(PowerManager.PARTIAL_WAKE_LOCK, "sanitas:ubicacion");
            wakeLock.acquire(12 * 60 * 60 * 1000L);
            registrarUbicacion();
            handler.post(ciclo);
        }
        return START_STICKY;
    }

    private void registrarUbicacion() {
        try {
            if (checkSelfPermission(android.Manifest.permission.ACCESS_FINE_LOCATION) != PackageManager.PERMISSION_GRANTED) return;
            lm = (LocationManager) getSystemService(Context.LOCATION_SERVICE);
            if (lm.isProviderEnabled(LocationManager.GPS_PROVIDER))
                lm.requestLocationUpdates(LocationManager.GPS_PROVIDER, 10000, 0, this, hilo.getLooper());
            if (lm.isProviderEnabled(LocationManager.NETWORK_PROVIDER))
                lm.requestLocationUpdates(LocationManager.NETWORK_PROVIDER, 10000, 0, this, hilo.getLooper());
        } catch (Exception ignorada) { }
    }

    @Override
    public void onLocationChanged(Location l) {
        Location previa = ultima;
        // Se prefiere el GPS; una ubicación de red solo reemplaza si la anterior ya es vieja o es menos precisa
        if (previa == null || LocationManager.GPS_PROVIDER.equals(l.getProvider())
                || System.currentTimeMillis() - previa.getTime() > 30000 || l.getAccuracy() <= previa.getAccuracy()) {
            ultima = l;
        }
    }
    @Override public void onProviderEnabled(String p) { }
    @Override public void onProviderDisabled(String p) { }
    @Override @SuppressWarnings("deprecation") public void onStatusChanged(String p, int s, Bundle e) { }

    // ───────────── ciclo de red ─────────────
    private String leerToken() {
        return getSharedPreferences(Nativo.PREFS, MODE_PRIVATE).getString(Nativo.K_TOKEN, null);
    }

    private void tick() throws Exception {
        SharedPreferences sp = getSharedPreferences(Nativo.PREFS, MODE_PRIVATE);
        String token = sp.getString(Nativo.K_TOKEN, null), prov = sp.getString(Nativo.K_PROV, null);
        if (token == null || prov == null) { detener(); return; }

        Location l = ultima;
        if (l != null && System.currentTimeMillis() - l.getTime() <= POSICION_MAX_EDAD_MS) {
            String body = String.format(Locale.US, "{\"lat\":%.7f,\"lng\":%.7f}", l.getLatitude(), l.getLongitude());
            int c = http("POST", "/proveedores/" + prov + "/ubicacion", token, body, null);
            if (c == 401 || c == 403) { sesionVencida(); return; }
        }

        StringBuilder resp = new StringBuilder();
        int c = http("GET", "/proveedores/" + prov + "/estado-actual", token, null, resp);
        if (c == 401 || c == 403) { sesionVencida(); return; }
        if (c != 200) return;

        JSONObject j = new JSONObject(resp.toString());
        String disp = j.optString("disponibilidad", "");
        if ("desconectado".equals(disp)) {            // el sistema lo desconectó (o lo hizo desde la app)
            avisoSimple("Te desconectamos", "Ya no recibes solicitudes. Abre la app y conéctate de nuevo.");
            detener();
            return;
        }
        JSONObject of = j.optJSONObject("oferta_pendiente");
        if (of != null) {
            long id = of.optLong("oferta_id", -1);
            if (id != ultimaOfertaAvisada) {
                ultimaOfertaAvisada = id;
                String tipo = "AMBULANCIA".equals(of.optString("tipo_servicio")) ? "Ambulancia" : "Médico a domicilio";
                String distrito = of.optString("distrito", "");
                avisoOferta("Nueva solicitud: " + tipo, (distrito.isEmpty() ? "Cerca de ti" : distrito) + ". Toca para responder.");
            }
        } else {
            NotificationManager nm = (NotificationManager) getSystemService(NOTIFICATION_SERVICE);
            nm.cancel(ID_OFERTA);                     // la oferta ya se respondió o venció
        }

        // Servicio asignado por la empresa: avisa una sola vez por solicitud
        JSONObject sv = j.optJSONObject("servicio_activo");
        if (sv != null && "asignado".equals(sv.optString("estado"))) {
            long sid = sv.optLong("solicitud_id", -1);
            if (sid != ultimoServicioAvisado) {
                ultimoServicioAvisado = sid;
                String tipo = "AMBULANCIA".equals(sv.optString("tipo_servicio")) ? "Ambulancia" : "Médico a domicilio";
                String distrito = sv.optString("distrito", "");
                avisoOferta("Te asignaron un servicio: " + tipo, (distrito.isEmpty() ? "" : distrito + ". ") + "Abre la app para ver los datos.");
            }
        }
    }

    private int http(String metodo, String ruta, String token, String body, StringBuilder salida) throws Exception {
        HttpURLConnection c = (HttpURLConnection) new URL(BuildConfig.BASE_URL + ruta).openConnection();
        try {
            c.setRequestMethod(metodo);
            c.setConnectTimeout(10000);
            c.setReadTimeout(10000);
            c.setRequestProperty("Authorization", "Bearer " + token);
            c.setRequestProperty("Content-Type", "application/json");
            if (body != null) {
                c.setDoOutput(true);
                try (OutputStream os = c.getOutputStream()) { os.write(body.getBytes(StandardCharsets.UTF_8)); }
            }
            int codigo = c.getResponseCode();
            if (salida != null && codigo == 200) {
                try (BufferedReader br = new BufferedReader(new InputStreamReader(c.getInputStream(), StandardCharsets.UTF_8))) {
                    String linea; while ((linea = br.readLine()) != null) salida.append(linea);
                }
            }
            return codigo;
        } catch (java.io.IOException sinRed) {
            return -1;                                // sin internet: se reintenta en el próximo ciclo
        } finally {
            c.disconnect();
        }
    }

    private void sesionVencida() {
        avisoSimple("Tu sesión venció", "Abre la app e inicia sesión para seguir recibiendo solicitudes.");
        detener();
    }

    private void detener() {
        getSharedPreferences(Nativo.PREFS, MODE_PRIVATE).edit().remove(Nativo.K_TOKEN).apply();
        stopSelf();
    }

    // ───────────── notificaciones ─────────────
    private void crearCanales() {
        if (Build.VERSION.SDK_INT < 26) return;
        NotificationManager nm = (NotificationManager) getSystemService(NOTIFICATION_SERVICE);
        NotificationChannel a = new NotificationChannel(CANAL_SERVICIO, "Ubicación activa", NotificationManager.IMPORTANCE_LOW);
        a.setDescription("Indica que la app está compartiendo tu ubicación mientras estás conectado");
        nm.createNotificationChannel(a);
        NotificationChannel b = new NotificationChannel(CANAL_OFERTAS, "Nuevas solicitudes", NotificationManager.IMPORTANCE_HIGH);
        b.enableVibration(true);
        b.setVibrationPattern(new long[]{0, 400, 200, 400, 200, 400});
        b.setLockscreenVisibility(Notification.VISIBILITY_PUBLIC);
        nm.createNotificationChannel(b);
    }

    private PendingIntent abrirApp() {
        Intent i = new Intent(this, MainActivity.class);
        i.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_SINGLE_TOP);
        return PendingIntent.getActivity(this, 0, i, PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
    }

    private Notification notificacionServicio(String texto) {
        Notification.Builder b = Build.VERSION.SDK_INT >= 26 ? new Notification.Builder(this, CANAL_SERVICIO) : new Notification.Builder(this);
        return b.setSmallIcon(R.drawable.ic_stat).setContentTitle("Sanitas Proveedor").setContentText(texto)
                .setOngoing(true).setContentIntent(abrirApp()).build();
    }

    @SuppressWarnings("deprecation")
    private void avisoOferta(String titulo, String texto) {
        Notification.Builder b = Build.VERSION.SDK_INT >= 26 ? new Notification.Builder(this, CANAL_OFERTAS) : new Notification.Builder(this);
        b.setSmallIcon(R.drawable.ic_stat).setContentTitle(titulo).setContentText(texto)
                .setCategory(Notification.CATEGORY_MESSAGE).setAutoCancel(true).setContentIntent(abrirApp())
                .setVisibility(Notification.VISIBILITY_PUBLIC);
        if (Build.VERSION.SDK_INT < 26) {
            b.setPriority(Notification.PRIORITY_HIGH).setDefaults(Notification.DEFAULT_ALL);
        }
        if (Build.VERSION.SDK_INT >= 26) b.setTimeoutAfter(120000);
        ((NotificationManager) getSystemService(NOTIFICATION_SERVICE)).notify(ID_OFERTA, b.build());
    }

    private void avisoSimple(String titulo, String texto) {
        Notification.Builder b = Build.VERSION.SDK_INT >= 26 ? new Notification.Builder(this, CANAL_OFERTAS) : new Notification.Builder(this);
        b.setSmallIcon(R.drawable.ic_stat).setContentTitle(titulo).setContentText(texto)
                .setAutoCancel(true).setContentIntent(abrirApp());
        ((NotificationManager) getSystemService(NOTIFICATION_SERVICE)).notify(3, b.build());
    }

    @Override
    public void onDestroy() {
        corriendo = false;
        try { if (lm != null) lm.removeUpdates(this); } catch (Exception ignorada) { }
        try { if (wakeLock != null && wakeLock.isHeld()) wakeLock.release(); } catch (Exception ignorada) { }
        if (hilo != null) hilo.quitSafely();
        super.onDestroy();
    }
}
