const d3 = require('d3-scale')
const Shell = require('node-powershell')

const ps = new Shell({
  executionPolicy: 'bypass',
  noProfile: true
})

go()

async function getVolume() {
  ps.addCommand('Get-AudioDevice -PlaybackVolume')
  const output = await ps.invoke()
  return parseInt(output.trim().replace('%', ''), 10)
}

async function setVolume(volume) {
  ps.addCommand(`Set-AudioDevice -PlaybackVolume ${volume}`)
  await ps.invoke()
}

async function getBrightness() {
  ps.addCommand(`(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness`)
  const output = await ps.invoke()
  return parseInt(output.trim(), 10)
}

async function setBrightness(percent) {
  ps.addCommand(`(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, ${percent})`)
  await ps.invoke()
}

async function suspendComputer() {
  ps.addCommand(`[Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")`)
  ps.addCommand(`$PowerState = [System.Windows.Forms.PowerState]::Suspend;`)
  ps.addCommand(`$Force = $false;`)
  ps.addCommand(`$DisableWake = $false;`)
  ps.addCommand(`[System.Windows.Forms.Application]::SetSuspendState($PowerState, $Force, $DisableWake);`)
  await ps.invoke()
}

function parseTime(timeish) {
  if (!timeish) {
    return 0
  }

  if (timeish.includes(':')) {
    const [mins, secs] = timeish.split(':').map(n => parseInt(n, 10))
    return mins * 60 * 1000 + secs * 1000
  } else {
    const secs = parseInt(timeish, 10)
    return secs * 1000
  }
}

async function go() {
  const duration = parseTime(process.argv[2])
  const delay = parseTime(process.argv[3])

  const volume = await getVolume()
  const brightness = await getBrightness()
  console.log(`Current volume is ${volume}`)
  console.log(`Current brightness is ${brightness}`)
  console.log(`Fading to 0 over ${duration} milliseconds`)

  if (delay > 0) {
    console.log(`but waiting ${delay} milliseconds before starting`)
    await new Promise(res => setTimeout(res, delay))
    console.log(`Starting fade...`)
  }

  const now = new Date().getTime()
  const volumeScale = d3.scaleLinear()
    .domain([now, now + duration])
    .range([volume, 0])
  const brightnessScale = d3.scaleLinear()
    .domain([now, now + duration])
    .range([brightness, 0])

  let lastVolume = volume
  let lastBrightness = brightness
  
  async function tick() {
    const newNow = new Date().getTime()
    const newVolume = volumeScale(newNow)
    const newBrightness = brightnessScale(newNow)

    if (newVolume <= 0) {
      await setVolume(0)
      await suspendComputer()
      process.exit(0)
    }

    if (Math.abs(lastVolume - newVolume) > 0.1) {
      const actualVolume = await getVolume()
      if (Math.abs(newVolume - actualVolume) > 0.1) {
        console.log(`Setting volume to ${newVolume}`)
        lastVolume = newVolume
        await setVolume(newVolume)
      }
    }

    if (Math.round(newBrightness) !== Math.round(lastBrightness)) {
      console.log(`Setting brightness to ${Math.round(newBrightness)}`)
      lastBrightness = newBrightness
      await setBrightness(Math.round(newBrightness))
    }

    process.nextTick(tick)
  }

  tick()
}