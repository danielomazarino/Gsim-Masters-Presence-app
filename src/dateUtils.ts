import { DateOption } from './types';

/**
 * Calculates relative date choices (Idag, Igår, 2 dagar sedan)
 * Formats dates for the backend as d/M (e.g. 4/9)
 */
export function calculateDateOptions(): DateOption[] {
  const now = new Date();

  function formatDate(d: Date): string {
    const day = d.getDate();
    const month = d.getMonth() + 1; // 1-12
    return `${day}/${month}`;
  }

  function formatDisplay(d: Date): string {
    const months = [
      'jan', 'feb', 'mar', 'apr', 'maj', 'jun',
      'jul', 'aug', 'sep', 'okt', 'nov', 'dec'
    ];
    const weekdays = ['Sön', 'Mån', 'Tis', 'Ons', 'Tor', 'Fre', 'Lör'];
    return `${weekdays[d.getDay()]} ${d.getDate()} ${months[d.getMonth()]}`;
  }

  const today = new Date(now);
  
  const yesterday = new Date(now);
  yesterday.setDate(now.getDate() - 1);

  const twoDaysAgo = new Date(now);
  twoDaysAgo.setDate(now.getDate() - 2);

  return [
    {
      id: 'today',
      label: 'Idag',
      dateFormatted: formatDate(today),
      subText: formatDisplay(today),
    },
    {
      id: 'yesterday',
      label: 'Igår',
      dateFormatted: formatDate(yesterday),
      subText: formatDisplay(yesterday),
    },
    {
      id: 'two_days_ago',
      label: '2 dagar sedan',
      dateFormatted: formatDate(twoDaysAgo),
      subText: formatDisplay(twoDaysAgo),
    },
  ];
}
