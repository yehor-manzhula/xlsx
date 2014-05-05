# Observable pattern
# @mixin
#

Observable =
  ###
  Subscribe
  @param {string} channel      event name
  @param {function} observer   function callback
  ###
  on: (channel, observer, once) ->
    @__channels = {}  unless @__channels
    @__channels[channel] = []  unless @__channels.hasOwnProperty(channel)
    @__channels[channel].push {observer:observer, once:once?}
    this

  once:(chanel, observer) ->
    @on chanel, observer, true

  ###
  Unsubscribe
  @param {string} channel      event name
  @param {function} observer   function callback
  ###
  off: (channel, observer) ->
    unless observer
      delete @__channels[channel]

      return this

    itemToRemove =  @__channels[channel].filter (element)->
      element.observer is observer

    @__channels[channel].splice @__channels[channel].indexOf(itemToRemove[0]), 1
    this

  ###
  Trigger event
  @param {string} channel      event name
  @param {function} data       data for called observers
  ###
  trigger: (channel, data) ->
    args = Array::slice.call arguments
    data = args.slice 1, args.length

    return this  if not @__channels or not @__channels[channel] or not @__channels[channel].length
    i = 0
    l = @__channels[channel].length

    while i < l
      currentElement = @__channels[channel][i]
      currentElement.observer.apply this, data

      if currentElement.once
        @off channel, currentElement.observer
        --l
      else
        ++i

    this

module.exports = Observable