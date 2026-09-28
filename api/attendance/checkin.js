import { supabase } from "../../lib/supabase";

import { calculateDistance } from "../../lib/geofence";

export default async function handler(req, res) {
  const { anggota_id, kegiatan_id, latitude, longitude } = req.body;

  const { data: geo } = await supabase
    .from("attendance_geofence_settings")
    .select("*")
    .eq("aktif", true)
    .single();

  const distance = calculateDistance(
    latitude,

    longitude,

    geo.latitude,

    geo.longitude,
  );

  if (distance > geo.radius_meter) {
    return res.json({
      success: false,

      message: "Anda berada di luar radius",

      distance,
    });
  }

  const { data, error } = await supabase
    .from("absensi")
    .insert({
      anggota_id,

      kegiatan_id,

      latitude,

      longitude,

      distance_meter: distance,

      status_kehadiran: "Hadir",

      jam_masuk: new Date(),
    })
    .select();

  if (error) throw error;

  res.json({
    success: true,

    data,
  });
}
