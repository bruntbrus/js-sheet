// Utility methods

(function() {
  'use strict'

  const utils = {

    isString(value) {
      return (typeof value === 'string')
    },

    isObject(value) {
      return (value && typeof value === 'object')
    },

    eachIn(object, callback) {
      Object.keys(object).forEach((key, i) => callback(key, object[key], i))
    },

    limit(num, min, max) {
      if (num < min) {
        num = min
      } else if (num > max) {
        num = max
      }

      return num
    },

    query(selector, root = document) {
      return root.querySelector(selector)
    },

    queryAll(selector, root = document) {
      return Array.from(root.querySelectorAll(selector))
    },

    text(string = '') {
      return document.createTextNode(string)
    },

    element(tagName, attributes = null) {
      const element = document.createElement(tagName)

      if (attributes) {
        if (typeof attributes === 'object') {
          utils.eachIn(attributes, (key, value) => {
            switch (key) {
              case '_text': element.textContent = value; break
              case '_html': element.innerHTML = value; break
              default: element.setAttribute(key, value)
            }
          })
        } else {
          element.textContent = attributes
        }
      }

      return element
    },

    getWidth(element) {
      return element.clientWidth
    },

    getHeight(element) {
      return element.clientHeight
    },

    setWidth(element, width) {
      element.style.width = width + 'px'
    },

    setHeight(element, height) {
      element.style.height = height + 'px'
    },

    setTop(element, top) {
      element.style.top = top + 'px'
    },

    setLeft(element, left) {
      element.style.left = left + 'px'
    },

    on(element, eventName, callback) {
      element.addEventListener(eventName, callback, false)
    },

    off(element, eventName, callback) {
      element.removeEventListener(eventName, callback, false)
    },

    debounce(delay, callback) {
      let timeout = false

      return function debounced(...args) {
        if (timeout !== false) {
          clearTimeout(timeout)
        }

        timeout = setTimeout(() => {
          timeout = false
          callback.apply(this, args)
        }, delay)
      }
    },

    round3(number) {
      return (Math.round(number * 1000) / 1000)
    },
  }

  // Export
  window.utils = utils
})()
