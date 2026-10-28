package pe.sanitas.proveedor;

import android.Manifest;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.provider.Settings;
import android.webkit.JavascriptInterface;

/**
 * Puente entre la página del proveedor (WebView) y el servicio nativo.
 * Se expone en JavaScript como window.MadNativo. Solo se carga la página del propio sistema.
 */
public class Nativo {
    static final String PREFS = "sanitas_proveedor";
    static final String K_TOKEN = "token";
    static final String K_PROV = "proveedor_id";

    private final Context ctx;

    public Nativo(Context ctx) {
        this.ctx = ctx.getApplicationContext();
    }

    /** Arranca (o actualiza) el servicio de ubicación. Devuelve "ok" o un código de falta de permiso. */
    @JavascriptInterface
    public String iniciar(String token, String proveedorId) {
        if (token == null || token.isEmpty() || proveedorId == null || !proveedorId.matches("\\d{1,10}")) {
            return "datos_invalidos";
        }
        if (ctx.checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) != PackageManager.PERMISSION_GRANTED) {
            return "falta_ubicacion";
        }
        SharedPreferences sp = ctx.getSharedPreferences(PREFS, Context.MODE_PRIVATE);
        sp.edit().putString(K_TOKEN, token).putString(K_PROV, proveedorId).apply();
        Intent i = new Intent(ctx, LocationService.class);
        try {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) ctx.startForegroundService(i);
            else ctx.startService(i);
        } catch (Exception e) {
            return "no_se_pudo_iniciar";
        }
        if (Build.VERSION.SDK_INT >= 33
                && ctx.checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) != PackageManager.PERMISSION_GRANTED) {
            return "falta_notificaciones";
        }
        return "ok";
    }

    @JavascriptInterface
    public void detener() {
        ctx.getSharedPreferences(PREFS, Context.MODE_PRIVATE).edit().remove(K_TOKEN).remove(K_PROV).apply();
        ctx.stopService(new Intent(ctx, LocationService.class));
    }

    @JavascriptInterface
    public String version() {
        return BuildConfig.VERSION_NAME;
    }

    @JavascriptInterface
    public void abrirAjustes() {
        Intent i = new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS, Uri.parse("package:" + ctx.getPackageName()));
        i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        ctx.startActivity(i);
    }
}
