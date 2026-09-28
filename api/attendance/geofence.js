import { supabase } from "../../lib/supabase";

export default async function handler(req, res) {
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

  res.json({
    success: true,

    data,
  });
}
