// Контакты компании. Пока поля пустые, блок "Связаться" на сайте скрыт.
// whatsapp и phone указывать цифрами с кодом страны, например '38267000000'. telegram: имя без @.
const CONTACT = { email: '', whatsapp: '', phone: '', telegram: '' };

// Приём заявок с формы "Связаться с нами". Пока endpoint пустой, форма показывает, что она не подключена.
// Подойдёт любой сервис, принимающий POST с JSON (например Formspree или Web3Forms).
// extra: дополнительные поля, которых требует сервис (например { access_key: '...' } для Web3Forms).
const FORM = { endpoint: '', extra: {} };

// Соцсети. Показываются только те, где указана ссылка (полный адрес, начиная с https://).
// Порядок на сайте: YouTube, Instagram, LinkedIn, Facebook.
// WhatsApp и Telegram здесь не нужны: они задаются выше, в CONTACT, и показываются в блоке "Быстрая связь".
const SOCIAL = {
  youtube: 'https://youtube.com/@allianzkapital',
  instagram: 'https://www.instagram.com/allianz_kapital/',
  linkedin: '',
  facebook: ''
};

// Headlines feed: refreshed every 30 min by .github/workflows/news.yml (needs that file on main).
// Falls back to the snapshot in docs/news.json.
const NEWS = { url: 'https://raw.githubusercontent.com/illyachaplygin/chaplygin-illya/news-data/news.json', pollMin: 5 };

// Счётчик посещений без cookie (баннер согласия не нужен). Пока пусто: счётчик выключен.
// goatcounter: код аккаунта с goatcounter.com (адрес вида КОД.goatcounter.com), plausible: домен сайта в plausible.io. Укажите что-то одно.
const ANALYTICS = { goatcounter: '', plausible: '' };
