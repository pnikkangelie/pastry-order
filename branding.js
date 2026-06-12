var DEFAULT_BRANDING = {
  name:       'Pastry Palette by Nikk',
  emoji:      '🥐',
  tagline:    'Every Dessert. A work of art',
  primary:    '#7b4a2c',
  accent:     '#c97d47',
  banner_img: ''
}

function getBranding() {
  try {
    var stored = JSON.parse(localStorage.getItem('scb_branding'))
    if (stored && Object.keys(stored).length) return stored
  } catch(e) {}
  // Fall back to bundled branding published by admin (for Netlify customers)
  if (typeof BUNDLED_BRANDING !== 'undefined' && BUNDLED_BRANDING && Object.keys(BUNDLED_BRANDING).length) {
    return BUNDLED_BRANDING
  }
  return DEFAULT_BRANDING
}

function tint(hex, ratio) {
  var n = parseInt(hex.slice(1), 16)
  var r = Math.round((n >> 16)       + (255 - (n >> 16))       * ratio)
  var g = Math.round(((n >> 8) & 255)+ (255 - ((n >> 8) & 255))* ratio)
  var b = Math.round((n & 255)       + (255 - (n & 255))       * ratio)
  return '#' + [r,g,b].map(function(v){ return v.toString(16).padStart(2,'0') }).join('')
}

function shade(hex, ratio) {
  var n = parseInt(hex.slice(1), 16)
  var r = Math.round((n >> 16)        * (1 - ratio))
  var g = Math.round(((n >> 8) & 255) * (1 - ratio))
  var b = Math.round((n & 255)        * (1 - ratio))
  return '#' + [r,g,b].map(function(v){ return v.toString(16).padStart(2,'0') }).join('')
}

function updateDocTitle(section) {
  var name = getBranding().name || 'My Bakery'
  document.title = section ? section + ' — ' + name : name
}

function applyBranding() {
  var b = getBranding()
  var root = document.documentElement
  root.style.setProperty('--primary',    b.primary)
  root.style.setProperty('--primary-h',  shade(b.primary, 0.15))
  root.style.setProperty('--primary-lt', tint(b.primary, 0.88))
  root.style.setProperty('--accent',     b.accent)
  document.querySelectorAll('.brand-emoji').forEach(function(el){ el.textContent = b.emoji || '' })
  document.querySelectorAll('.brand-name').forEach(function(el){ el.textContent = b.name || 'My Bakery' })
  // Page background color — reset to CSS default (#faf8f5) when not set
  root.style.setProperty('--bg', b.bg_color || '#faf8f5')
  // Banner image — allow http/https URLs and data:image/ URIs
  var hero = document.getElementById('order-hero')
  if (hero) {
    var bannerUrl = (b.banner_img && (/^https?:\/\//i.test(b.banner_img) || /^data:image\//i.test(b.banner_img))) ? b.banner_img : ''
    if (bannerUrl) {
      hero.style.backgroundImage = 'url("' + bannerUrl.replace(/"/g, '%22') + '")'
      hero.classList.add('has-banner-img')
    } else {
      hero.style.backgroundImage = ''
      hero.classList.remove('has-banner-img')
    }
  }
  // Tagline (order form hero only)
  var tagline = document.getElementById('hero-tagline')
  if (tagline && b.tagline !== undefined) tagline.textContent = b.tagline || ''
}
