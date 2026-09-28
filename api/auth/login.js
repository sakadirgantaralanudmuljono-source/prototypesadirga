import { supabase } from "../../lib/supabase";

export default async function handler(req, res) {
  if (req.method !== "POST")
    return res.status(405).json({
      error: "Method not allowed",
    });

  const { email, password } = req.body;

  const { data, error } = await supabase.auth.signInWithPassword({
    email,

    password,
  });

  if (error)
    return res.status(401).json({
      success: false,

      message: error.message,
    });

  return res.json({
    success: true,

    session: data.session,

    user: data.user,
  });
}
