package pe.sanitas.proveedor;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.PowerManager;
import android.provider.Settings;
import android.webkit.GeolocationPermissions;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import java.util.ArrayList;
import java.util.List;

public class MainActivity extends Activity {
    private static final int REQ_PERMISOS = 10;
    private WebView web;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        web = new WebView(this);
        setContentView(web);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setGeolocationEnabled(true);
        s.setCacheMode(WebSettings.LOAD_DEFAULT);
        s.setUserAgentString(s.getUserAgentString() + " SanitasProveedorApp/" + BuildConfig.VERSION_NAME);

        web.addJavascriptInterface(new Nativo(this), "MadNativo");

        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onGeolocationPermissionsShowPrompt(String origin, GeolocationPermissions.Callback cb) {
                boolean ok = checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) == PackageManager.PERMISSION_GRANTED;
                cb.invoke(origin, ok, false);
            }
        });

        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest req) {
                Uri u = req.getUrl();
                String base = Uri.parse(BuildConfig.BASE_URL).getHost();
                if ("https".equals(u.getScheme()) && base != null && base.equals(u.getHost())) return false;
                // Llamadas (tel:), WhatsApp y cualquier otro sitio se abren fuera de la app
                try {
                    Intent i = new Intent("tel".equals(u.getScheme()) ? Intent.ACTION_DIAL : Intent.ACTION_VIEW, u);
                    startActivity(i);
                } catch (Exception ignorada) { }
                return true;
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest req, WebResourceError err) {
                if (!req.isForMainFrame()) return;
                String html = "<html><head><meta name='viewport' content='width=device-width,initial-scale=1'></head>"
                        + "<body style='font-family:sans-serif;text-align:center;padding:48px 24px;color:#003876'>"
                        + "<h2>Sin conexión</h2><p>Revisa tu internet e intenta de nuevo.</p>"
                        + "<button style='font-size:18px;padding:12px 28px;border:0;border-radius:10px;background:#003876;color:#fff' "
                        + "onclick=\"location.href='" + BuildConfig.BASE_URL + "/proveedor'\">Reintentar</button></body></html>";
                view.loadDataWithBaseURL(null, html, "text/html", "UTF-8", null);
            }
        });

        if (savedInstanceState == null) web.loadUrl(BuildConfig.BASE_URL + "/proveedor");
        pedirPermisos();
    }

    private void pedirPermisos() {
        List<String> faltan = new ArrayList<>();
        if (checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) != PackageManager.PERMISSION_GRANTED) {
            faltan.add(Manifest.permission.ACCESS_FINE_LOCATION);
            faltan.add(Manifest.permission.ACCESS_COARSE_LOCATION);
        }
        if (Build.VERSION.SDK_INT >= 33
                && checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) != PackageManager.PERMISSION_GRANTED) {
            faltan.add(Manifest.permission.POST_NOTIFICATIONS);
        }
        if (!faltan.isEmpty()) {
            requestPermissions(faltan.toArray(new String[0]), REQ_PERMISOS);
        } else {
            pedirSinRestriccionDeBateria();
        }
    }

    @Override
    public void onRequestPermissionsResult(int code, String[] perms, int[] results) {
        super.onRequestPermissionsResult(code, perms, results);
        if (code == REQ_PERMISOS) {
            web.reload();                     // la página vuelve a pedir iniciar el servicio con el permiso ya concedido
            pedirSinRestriccionDeBateria();
        }
    }

    /** Sin esto muchos celulares (Xiaomi, Samsung, Huawei...) cierran el servicio con la pantalla bloqueada. */
    private void pedirSinRestriccionDeBateria() {
        try {
            SharedPreferences sp = getSharedPreferences(Nativo.PREFS, Context.MODE_PRIVATE);
            PowerManager pm = (PowerManager) getSystemService(Context.POWER_SERVICE);
            if (sp.getBoolean("bateria_pedida", false) || pm.isIgnoringBatteryOptimizations(getPackageName())) return;
            sp.edit().putBoolean("bateria_pedida", true).apply();
            Intent i = new Intent(Settings.ACTION_REQUEST_IGNORE_BATTERY_OPTIMIZATIONS, Uri.parse("package:" + getPackageName()));
            startActivity(i);
        } catch (Exception ignorada) { }
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        out.putBoolean("cargada", true);
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else moveTaskToBack(true);           // no cierra la app: sigue compartiendo la ubicación
    }

    @Override
    protected void onResume() {
        super.onResume();
        web.onResume();
    }
}
