/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./node_modules/flowbite/**/*.js",
    './index.html', "sucess.html"],
  theme: {
    extend: {},
    colors: {
      tittle: '#b91c1c',
      textblack: '#27272a',
      alert: '#FF0000',
    },
  },
  plugins: [
    require('flowbite/plugin'),
    require('daisyui'),
  ],
  
}

