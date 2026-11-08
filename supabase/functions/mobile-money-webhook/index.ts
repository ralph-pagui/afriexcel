// Supabase Edge Function: mobile-money-webhook
// Réception des webhooks des passerelles de paiement (CinetPay, Wave, FedaPay, PayTech)
// Valide le paiement et crédite instantanément le compte de l'utilisateur

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.49.1";

const supabase = createClient(
  Deno.env.get("SUPABASE_URL") ?? "",
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? ""
);

serve(async (req) => {
  try {
    const payload = await req.json();
    console.log("Notification Mobile Money reçue :", payload);

    // Extraction des détails selon la passerelle (CinetPay / Wave / FedaPay)
    const transactionId = payload.cpm_trans_id || payload.transaction_id || payload.id;
    const status = payload.cpm_result === "00" || payload.status === "approved" || payload.status === "SUCCESS" 
      ? "SUCCESS" 
      : "FAILED";

    if (transactionId && status === "SUCCESS") {
      // 1. Mettre à jour la transaction
      const { data: tx } = await supabase
        .from("mobile_money_transactions")
        .update({ status: "SUCCESS" })
        .eq("transaction_ref", transactionId)
        .select()
        .single();

      // 2. Activer l'abonnement ou ajouter des crédits
      if (tx && tx.user_id) {
        await supabase
          .from("subscriptions")
          .update({
            is_active: true,
            plan: "pro_mensuel",
            credits_remaining: 500,
            expires_at: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString()
          })
          .eq("user_id", tx.user_id);
      }
    }

    return new Response(JSON.stringify({ received: true }), {
      headers: { "Content-Type": "application/json" },
      status: 200
    });
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500
    });
  }
});
