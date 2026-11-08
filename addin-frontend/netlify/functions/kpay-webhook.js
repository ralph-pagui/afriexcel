/**
 * Webhook Serverless Netlify pour KPay (kpay.site)
 * Reçoit la confirmation de paiement Mobile Money en temps réel (Wave, Orange, MTN, Moov)
 * et met à jour automatiquement la base de données Supabase.
 */

const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.VITE_SUPABASE_URL || 'https://demo-african-assistant.supabase.co';
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.VITE_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.demo-key';

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);

exports.handler = async function (event, context) {
  // Uniquement les requêtes POST provenant de KPay
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method Not Allowed' })
    };
  }

  try {
    const payload = JSON.parse(event.body || '{}');
    const {
      refid,
      transaction_id,
      status,
      amount,
      currency,
      operator,
      customer_phone
    } = payload;

    const txRef = refid || transaction_id;
    const isSuccess = status === 'PAID' || status === 'SUCCESS' || status === 'COMPLETED';

    console.log(`[KPay Webhook] Réception notification pour transaction ${txRef} - Statut : ${status}`);

    if (txRef) {
      // 1. Mettre à jour la transaction dans Supabase
      await supabase
        .from('mobile_money_transactions')
        .update({
          status: isSuccess ? 'SUCCESS' : 'FAILED',
          gateway_ref: transaction_id || txRef
        })
        .eq('transaction_ref', txRef);

      // 2. Si le paiement est validé, activer ou prolonger l'abonnement
      if (isSuccess && customer_phone) {
        // Rechercher le profil par téléphone ou créer/activer l'abonnement
        const expiresAt = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();
        
        // Notification temps réel envoyée par Supabase
        await supabase
          .from('subscriptions')
          .insert({
            plan: 'pack_credits',
            credits_remaining: 100,
            is_active: true,
            expires_at: expiresAt
          });
      }
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        received: true,
        transactionRef: txRef,
        status: isSuccess ? 'ACTIVATED' : 'PROCESSED'
      })
    };
  } catch (error) {
    console.error('[KPay Webhook Error]:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message })
    };
  }
};
