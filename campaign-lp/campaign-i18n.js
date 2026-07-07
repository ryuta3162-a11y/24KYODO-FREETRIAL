/**
 * キャンペーンLP — JSで差し替わる文言（言語切替時に再適用）
 */
(function () {
  function t(key, fallback) {
    return window.JoyfitI18n ? JoyfitI18n.t(key, fallback) : fallback;
  }

  function setText(id, key, fallback) {
    var el = document.getElementById(id);
    if (el) el.textContent = t(key, fallback);
  }

  function setHtml(id, key, fallback) {
    var el = document.getElementById(id);
    if (el) el.innerHTML = t(key, fallback);
  }

  function applyCampaignDynamic() {
    setText('campaign-option-label', 'pricing.optionLabel', '7月オプション');
    setText('campaign-option-label-sub', 'pricing.optionLabelSub', '無料オプション8つ自動契約');
    setText('campaign-option-highlight', 'pricing.optionHighlight', '7月分が0円');
    setHtml(
      'campaign-option-note',
      'pricing.optionNote',
      '※8つのオプションは入会時自動契約されます。<br>不要な場合は<span id="campaign-option-cancel-deadline">7月末</span>までに解約をお願いいたします。'
    );
    setText('campaign-option-cancel-deadline', 'pricing.optionCancelDeadline', '7月末');
    if (window.JoyfitI18n && JoyfitI18n.getLanguage() !== 'ja') {
      setText('dynamic-date-label', 'pricing.joinAmountLabel', 'ご入会時金額');
    }
    if (typeof window.updateProratedFee === 'function') {
      window.updateProratedFee();
    }
  }

  document.addEventListener('DOMContentLoaded', applyCampaignDynamic);
  window.addEventListener('joyfit:langchange', applyCampaignDynamic);
})();
