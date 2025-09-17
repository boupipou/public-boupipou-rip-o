export async function injectFooter(path = 'footer.html') {
  try {
    const html = await fetch(path).then(r => r.text());
    document.body.insertAdjacentHTML('beforeend', html);
  } catch (err) {
    console.error('Footer failed to load:', err);
  }
}