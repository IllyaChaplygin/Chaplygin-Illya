// Контакты компании. Пока поля пустые, блок "Связаться" на сайте скрыт.
// whatsapp и phone указывать цифрами с кодом страны, например '38267000000'. telegram: имя без @.
const CONTACT = { email: '', whatsapp: '', phone: '', telegram: '' };

// Приём заявок с формы "Связаться с нами". Пока endpoint пустой, форма показывает, что она не подключена.
// Подойдёт любой сервис, принимающий POST с JSON (например Formspree или Web3Forms).
// extra: дополнительные поля, которых требует сервис (например { access_key: '...' } для Web3Forms).
const FORM = { endpoint: '', extra: {} };

// Соцсети. Показываются только те, где указана ссылка (полный адрес, начиная с https://).
// Порядок на сайте: YouTube, Instagram, LinkedIn, Facebook, Telegram, TikTok, Google.
const SOCIAL = {
  youtube: 'https://youtube.com/@allianzkapital',
  instagram: '',
  linkedin: '',
  facebook: '',
  telegram: '',   // канал или группа, например https://t.me/название
  tiktok: '',
  google: ''      // страница компании в Google Картах (отзывы)
};
