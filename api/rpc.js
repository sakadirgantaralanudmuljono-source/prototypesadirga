import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
);

const SUPABASE_METHODS = new Set([
  "getAttendanceGeofenceSettings",
  "saveAttendanceGeofenceSettings",
  "login",
  "logout",
  "getCurrentUser",
]);

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      success: false,
      message: "Method not allowed",
    });
  }

  const body = req.body || {};
  const method = String(body.method || "").trim();
  const args = Array.isArray(body.args) ? body.args : [];

  if (SUPABASE_METHODS.has(method)) {
    return handleSupabaseMethod(method, args, res);
  }

  return res.status(500).json({
    success: false,
    message: "Method belum dimigrasi ke Supabase: " + method,
  });
}

async function handleSupabaseMethod(method, args, res) {
  switch (method) {
    case "getAttendanceGeofenceSettings":
      const { data, error } = await supabase
        .from("attendance_geofence_settings")
        .select("*")
        .eq("aktif", true);

      if (error) {
        return res.status(500).json({
          success: false,
          message: error.message,
        });
      }

      return res.json({
        success: true,
        data,
      });

    case "login":
      const payload = args[0] || {};

      const { data, error: loginError } =
        await supabase.auth.signInWithPassword({
          email: payload.email,
          password: payload.password,
        });

      if (loginError) {
        return res.status(401).json({
          success: false,
          message: loginError.message,
        });
      }

      const { data: profile } = await supabase
        .from("profiles")
        .select("*")
        .eq("id", data.user.id)
        .single();

      return res.json({
        success: true,
        data: {
          session: data.session,
          user: data.user,
          profile,
        },
      });

    case "logout":
      await supabase.auth.signOut();

      return res.json({
        success: true,
      });

    default:
      return res.status(400).json({
        success: false,
        message: "Supabase method tidak tersedia",
      });
  }
}
