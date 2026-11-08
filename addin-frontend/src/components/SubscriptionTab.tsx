import React, { useState } from 'react';
import { CreditCard, Smartphone, CheckCircle, Zap, ShieldCheck, ArrowRight, Loader2 } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';
import { PaymentService, PaymentRequest } from '../services/paymentService';

interface SubscriptionTabProps {
  country: AfricanCountry;
}

export const SubscriptionTab: React.FC<SubscriptionTabProps> = ({ country }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [selectedPlan, setSelectedPlan] = useState<'MONTHLY' | 'PACK'>('MONTHLY');
  const [operator, setOperator] = useState<'WAVE' | 'ORANGE_MONEY' | 'MTN_MOMO' | 'MOOV_MONEY' | 'CARD'>('WAVE');
  const [phone, setPhone] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [paymentSuccess, setPaymentSuccess] = useState<string | null>(null);

  const planAmount = selectedPlan === 'MONTHLY' ? 5000 : 2500;
  const planLabel = selectedPlan === 'MONTHLY' ? 'Abonnement Pro Mensuel (Illimité)' : 'Pack 100 Crédits d\'analyses';

  const handlePay = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsProcessing(true);
    setPaymentSuccess(null);

    try {
      const res = await PaymentService.initiateMobileMoneyPayment({
        operator,
        phoneNumber: phone,
        amount: planAmount,
        currency: countryInfo.currencySymbol,
        planName: planLabel
      });

      if (res.success) {
        setPaymentSuccess(res.message);
      } else {
        alert(res.message);
      }
    } catch (err: any) {
      alert("Erreur de paiement : " + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      {/* En-tête */}
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <Smartphone className="w-4 h-4 text-emerald-600" />
          <span>Paiement Sans Friction Mobile Money</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Activez votre accès en 30 secondes via <strong>Wave, Orange Money, MTN MoMo</strong> ou carte bancaire.
        </p>
      </div>

      {/* Choix de la formule */}
      <div className="space-y-2">
        <label className="text-xs font-semibold text-slate-700">Choisissez votre formule :</label>
        <div className="grid grid-cols-2 gap-2">
          <button
            type="button"
            onClick={() => setSelectedPlan('MONTHLY')}
            className={`p-3 rounded-xl border text-left transition-all ${
              selectedPlan === 'MONTHLY'
                ? 'border-emerald-600 bg-emerald-50/50 ring-2 ring-emerald-500'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-700 bg-emerald-100 px-1.5 py-0.5 rounded">Recommandé</span>
              <Zap className="w-3.5 h-3.5 text-amber-500 fill-amber-500" />
            </div>
            <div className="font-bold text-slate-900 text-xs">Pro Mensuel</div>
            <div className="text-sm font-extrabold text-emerald-700 mt-0.5">5 000 {countryInfo.currencySymbol}</div>
            <div className="text-[10px] text-slate-500 mt-1">Usage illimité</div>
          </button>

          <button
            type="button"
            onClick={() => setSelectedPlan('PACK')}
            className={`p-3 rounded-xl border text-left transition-all ${
              selectedPlan === 'PACK'
                ? 'border-emerald-600 bg-emerald-50/50 ring-2 ring-emerald-500'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between mb-1">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-600 bg-slate-100 px-1.5 py-0.5 rounded">À la carte</span>
            </div>
            <div className="font-bold text-slate-900 text-xs">Pack Crédits</div>
            <div className="text-sm font-extrabold text-slate-800 mt-0.5">2 500 {countryInfo.currencySymbol}</div>
            <div className="text-[10px] text-slate-500 mt-1">100 requêtes valables 1 an</div>
          </button>
        </div>
      </div>

      {/* Formulaire de paiement */}
      <form onSubmit={handlePay} className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-3 text-xs">
        <div>
          <label className="font-semibold text-slate-700 block mb-1.5">Sélectionnez votre moyen de paiement :</label>
          <div className="grid grid-cols-2 gap-1.5">
            {[
              { id: 'WAVE', label: 'Wave', color: 'text-sky-500 font-bold' },
              { id: 'ORANGE_MONEY', label: 'Orange Money', color: 'text-orange-500 font-bold' },
              { id: 'MTN_MOMO', label: 'MTN MoMo', color: 'text-yellow-600 font-bold' },
              { id: 'MOOV_MONEY', label: 'Moov Money', color: 'text-blue-600 font-bold' }
            ].map(op => (
              <button
                key={op.id}
                type="button"
                onClick={() => setOperator(op.id as any)}
                className={`py-2 px-2.5 rounded-lg border text-left flex items-center justify-between text-xs transition-all ${
                  operator === op.id
                    ? 'border-emerald-600 bg-emerald-50 text-slate-900 font-medium'
                    : 'border-slate-200 hover:bg-slate-50 text-slate-700'
                }`}
              >
                <span className={op.color}>{op.label}</span>
                {operator === op.id && <CheckCircle className="w-3.5 h-3.5 text-emerald-600" />}
              </button>
            ))}
          </div>
        </div>

        <div>
          <label className="font-semibold text-slate-700 block mb-1">
            Numéro {operator.replace('_', ' ')} :
          </label>
          <div className="flex items-center gap-1.5">
            <span className="p-2 bg-slate-100 border border-slate-300 rounded-lg text-slate-600 font-medium shrink-0">
              {countryInfo.flag}
            </span>
            <input
              type="tel"
              required
              value={phone}
              onChange={e => setPhone(e.target.value)}
              placeholder="Ex: 07 08 09 10 11"
              className="flex-1 border border-slate-300 rounded-lg px-3 py-2 text-xs focus:ring-2 focus:ring-emerald-500 outline-none font-medium"
            />
          </div>
        </div>

        <button
          type="submit"
          disabled={isProcessing || !phone.trim()}
          className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-bold py-3 px-4 rounded-xl text-xs shadow-md transition-all"
        >
          {isProcessing ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin" />
              Validation du paiement en cours...
            </>
          ) : (
            <>
              <span>Payer {planAmount} {countryInfo.currencySymbol}</span>
              <ArrowRight className="w-4 h-4" />
            </>
          )}
        </button>
      </form>

      {paymentSuccess && (
        <div className="p-3.5 bg-emerald-100 border border-emerald-300 rounded-xl text-emerald-900 text-xs space-y-1">
          <div className="flex items-center gap-1.5 font-bold">
            <CheckCircle className="w-4 h-4 text-emerald-700" />
            <span>Paiement initié !</span>
          </div>
          <p className="text-[11px] text-emerald-800">{paymentSuccess}</p>
        </div>
      )}

      <div className="flex items-center justify-center gap-2 text-[10px] text-slate-500 pt-1">
        <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
        <span>Paiement sécurisé et instantané via passerelle Mobile Money certifiée BCEAO.</span>
      </div>
    </div>
  );
};
