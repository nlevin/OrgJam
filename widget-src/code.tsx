const { widget } = figma
const { usePropertyMenu, useSyncedState, useWidgetNodeId, waitForTask, AutoLayout, Image, Input, Text, SVG } = widget

type WidgetKind = 'chooser' | 'people-chip' | 'counter' | 'bulk-create'

type Status = 'active' | 'open-headcount' | 'new-headcount'
type CounterPreset = 'all' | 'locations' | Status
type ManagerKind = 'ic' | 'manager'
type CounterPeopleFilter = 'all' | 'ics' | 'managers'
type ChipProfile = {
  name: string
  projectArea: string
  role: string
  location: string
  status: Status
}
type ParsedCardCandidate = {
  sourceNodeId: string
  x: number
  y: number
  width: number
  height: number
  profile: ChipProfile
}
type ParsedConnection = {
  fromSourceId: string
  toSourceId: string
}
type ImportParseResult = {
  cards: ParsedCardCandidate[]
  connections: ParsedConnection[]
  containsRasterImage: boolean
}
type BulkSource = 'text' | 'selection'
type IndexedImportNode = {
  id: string
  profile: ChipProfile
  order: number
}
type CounterRoleDetail = {
  title: string
  color: string
  statusSummary: string
}
type OcrOverlayLine = {
  text: string
  left: number
  top: number
  width: number
  height: number
}
type AvatarUploadMessage =
  | { type: 'avatar-uploaded'; dataUri: string }
  | { type: 'avatar-upload-error'; message: string }
  | { type: 'avatar-remove-image' }
type LinkEditMessage =
  | { type: 'link-save'; url: string }
  | { type: 'link-cancel' }

const PEOPLE_CHIP_WIDTH = 260
const PEOPLE_CHIP_DETAILS_WIDTH = 164
const COUNTER_DETAILS_WIDTH = 162
const CHOOSER_PEOPLE_CARD_WIDTH = PEOPLE_CHIP_WIDTH
const CHOOSER_SMALL_CARD_WIDTH = 110
const CHOOSER_CARD_HEIGHT = 80
const OCR_SPACE_ENDPOINT = 'https://api.ocr.space/parse/image'
const OCR_SPACE_API_KEY = 'K83280778188957'
const AVATAR_UPLOAD_TIMEOUT_MS = 3 * 60 * 1000

const STATUS_OPTIONS: Array<{ option: Status; label: string }> = [
  { option: 'active', label: 'Active' },
  { option: 'open-headcount', label: 'Open' },
  { option: 'new-headcount', label: 'New' },
]

const ROLE_OPTIONS: Array<{ option: string; label: string }> = [
  { option: 'Design', label: 'Design' },
  { option: 'Eng', label: 'Eng' },
  { option: 'PM', label: 'PM' },
  { option: 'TPM', label: 'TPM' },
  { option: 'Ops', label: 'Ops' },
  { option: 'Brand', label: 'Brand' },
  { option: 'Writing', label: 'Writing' },
  { option: 'Research', label: 'Research' },
  { option: 'Data', label: 'Data' },
]

const COUNTER_PRESET_OPTIONS: Array<{ option: CounterPreset; label: string }> = [
  { option: 'all', label: 'Total' },
  { option: 'active', label: 'Active' },
  { option: 'open-headcount', label: 'Open' },
  { option: 'new-headcount', label: 'New' },
  { option: 'locations', label: 'Locations' },
]

const COUNTER_PRESET_LABEL: Record<CounterPreset, string> = {
  all: 'Total',
  locations: 'Locations',
  active: 'Active',
  'open-headcount': 'Open',
  'new-headcount': 'New',
}

const COUNTER_PEOPLE_FILTER_OPTIONS: Array<{ option: CounterPeopleFilter; label: string }> = [
  { option: 'all', label: 'All' },
  { option: 'ics', label: 'ICs' },
  { option: 'managers', label: 'Managers' },
]

const STATUS_STYLES: Record<
  Status,
  {
    label: string
    cardFill: string
    cardStroke: string
    avatarFill: string
    avatarStroke: string
    avatarText: string
  }
> = {
  active: {
    label: 'Active',
    cardFill: '#FFFFFF',
    cardStroke: '#CFCFD4',
    avatarFill: '#F4F4F5',
    avatarStroke: '#D4D4D8',
    avatarText: '#6B7280',
  },
  'open-headcount': {
    label: 'Open',
    cardFill: '#F5F5F5',
    cardStroke: '#CFCFD4',
    avatarFill: '#E8E8E8',
    avatarStroke: '#D4D4D8',
    avatarText: '#6B7280',
  },
  'new-headcount': {
    label: 'New',
    cardFill: '#FFF6DC',
    cardStroke: '#FFE8A3',
    avatarFill: '#FFEFC1',
    avatarStroke: '#E9D899',
    avatarText: '#CDAD4D',
  },
}

const DEFAULT_CHIP_PROFILES = [
  { name: 'Alex Kim', projectArea: 'Search', role: 'Designer', location: 'NYC' },
  { name: 'Priya Patel', projectArea: 'Platform', role: 'Eng', location: 'San Francisco' },
  { name: 'Jordan Lee', projectArea: 'Infra', role: 'PM', location: 'London' },
  { name: 'Sam Rivera', projectArea: 'Growth', role: 'Eng', location: 'Austin' },
  { name: 'Maya Chen', projectArea: 'AI', role: 'Designer', location: 'Seattle' },
  { name: 'Diego Alvarez', projectArea: 'Security', role: 'PM', location: 'Toronto' },
] as const

type Point = { x: number; y: number }

function refreshIconSvg(fill: string): string {
  return `<svg width="22" height="15" viewBox="0 0 22 15" fill="none" xmlns="http://www.w3.org/2000/svg">
  <path fill-rule="evenodd" clip-rule="evenodd" d="M11.9026 1.43168C12.1936 1.47564 12.4822 1.54098 12.7663 1.62777L12.7719 1.62949C14.0176 2.0114 15.109 2.78567 15.8858 3.83854L15.8918 3.84665C16.5473 4.73808 16.9484 5.78867 17.058 6.88508L14.0863 4.88858L13.3259 6.02047L17.3852 8.74774L17.9079 9.09894L18.2994 8.60571L21.0056 5.19662L19.9376 4.34879L18.3531 6.34479C18.3424 6.27511 18.3306 6.20563 18.3179 6.13636C18.1135 5.02233 17.6601 3.96334 16.9851 3.04274L16.9791 3.03462C16.0303 1.74427 14.6956 0.794984 13.1714 0.326388L13.1658 0.32466C12.8171 0.217755 12.4627 0.137298 12.1055 0.0832198C10.899 -0.0994351 9.66061 0.0188515 8.50099 0.435448L8.4947 0.437711C7.42511 0.823053 6.46311 1.44778 5.6774 2.25801C5.38576 2.55876 5.11841 2.88506 4.87886 3.23416C4.85856 3.26376 4.83845 3.29351 4.81854 3.32343L5.94262 4.08294L5.94802 4.07484C5.96253 4.0531 5.97717 4.03146 5.99195 4.00993C6.71697 2.95331 7.75331 2.15199 8.95541 1.72013L8.9617 1.71788C9.33245 1.58514 9.71301 1.48966 10.098 1.43156C10.6957 1.34135 11.3039 1.34123 11.9026 1.43168ZM3.70034 6.39429L0.994141 9.80338L2.06217 10.6512L3.64663 8.65521C3.65741 8.72489 3.66916 8.79437 3.68187 8.86364C3.88627 9.97767 4.33964 11.0367 5.01467 11.9573L5.02063 11.9654C5.96945 13.2557 7.30418 14.205 8.82835 14.6736L8.83398 14.6753C9.18281 14.7823 9.53732 14.8628 9.89464 14.9168C11.101 15.0994 12.3393 14.9811 13.4988 14.5646L13.5051 14.5623C14.5747 14.1769 15.5367 13.5522 16.3224 12.742C16.614 12.4413 16.8813 12.115 17.1209 11.7659C17.1412 11.7363 17.1613 11.7065 17.1812 11.6766L16.0571 10.9171L16.0518 10.9252C16.0372 10.9469 16.0225 10.9686 16.0078 10.9902C15.2827 12.0467 14.2464 12.848 13.0444 13.2799L13.0381 13.2821C12.6673 13.4149 12.2868 13.5103 11.9018 13.5684C11.3041 13.6587 10.6958 13.6588 10.0971 13.5683C9.8062 13.5244 9.51754 13.459 9.23347 13.3722L9.22784 13.3705C7.98212 12.9886 6.89078 12.2143 6.11393 11.1615L6.10795 11.1534C5.45247 10.2619 5.05138 9.21133 4.94181 8.11492L7.91342 10.1114L8.6739 8.97953L4.61459 6.25226L4.09188 5.90106L3.70034 6.39429Z" fill="${fill}"/>
</svg>`
}

const TOOLBAR_ICON_COLOR = '#A7A8AE'
const LINK_ICON = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M3.12246 5.35035C3.70967 5.29718 4.16535 5.85058 3.84109 6.39269C3.68701 6.63352 3.44221 6.83986 3.24592 7.04265C2.35156 7.96672 1.28591 8.70535 1.39299 10.1548C1.44415 10.8475 1.7144 11.4879 2.23257 11.958C2.64677 12.3278 3.16901 12.5519 3.7206 12.5966C6.06456 12.7748 6.99864 10.1048 7.94124 10.0672C8.11707 10.0602 8.28682 10.1478 8.41128 10.2682C8.55093 10.4033 8.64159 10.593 8.64047 10.79C8.63868 11.0981 8.41314 11.302 8.21515 11.5043C7.53853 12.1888 6.73037 13.1189 5.89479 13.5403C4.875 14.0543 3.69489 14.141 2.61186 13.7816C1.65908 13.4519 0.874374 12.7557 0.428993 11.8452C-0.0572896 10.8603 -0.133474 9.72056 0.217403 8.67886C0.33225 8.32873 0.498941 7.99808 0.711947 7.69823C1.02018 7.26336 2.62075 5.57856 3.06514 5.35706C3.08414 5.35422 3.10331 5.35194 3.12246 5.35035ZM9.50159 3.64786C9.85021 3.62117 10.1332 3.80056 10.2804 4.12307C10.3339 4.22013 10.3481 4.51678 10.2816 4.60991C9.92367 5.11129 9.29651 5.6841 8.87419 6.10817L5.6311 9.37951C5.31915 9.69468 4.90772 10.1883 4.50274 10.3454C3.99895 10.4182 3.4979 9.88582 3.72286 9.38802C3.76671 9.29096 3.96731 9.06785 4.04257 8.98963C4.38325 8.63561 4.7349 8.29059 5.08114 7.94189L6.96904 6.03937C7.63591 5.36991 8.29662 4.69294 8.96791 4.02803C9.13393 3.8636 9.28283 3.73463 9.50159 3.64786ZM9.80749 0.0145755C10.4524 -0.0563199 11.2955 0.139268 11.8723 0.430699C12.7799 0.898545 13.4677 1.70851 13.7859 2.6845C14.1328 3.71924 14.0558 4.85076 13.5721 5.8282C13.0931 6.77384 11.9844 7.65572 11.2545 8.43436C10.8622 8.85291 10.1616 8.63147 10.0533 8.06018C9.97124 7.62683 10.4734 7.26495 10.7319 6.98116C11.2481 6.42485 11.9088 5.91491 12.2895 5.25495C12.881 4.15661 12.6525 2.76563 11.6774 1.97306C11.2462 1.62246 10.7135 1.42259 10.1597 1.40353C7.89993 1.32591 7.02925 3.83633 6.07332 3.92855C5.91455 3.94385 5.77037 3.89138 5.64883 3.78941C5.48655 3.65325 5.36455 3.47564 5.34886 3.25871C5.33349 3.04628 5.43901 2.8681 5.57271 2.71338C5.84333 2.40021 6.15684 2.11149 6.44833 1.8175C7.49176 0.765096 8.26819 0.12648 9.80749 0.0145755Z" fill="${TOOLBAR_ICON_COLOR}"/>
</svg>`
const OPEN_LINK_ICON = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M3.16338 0.0172543C3.64039 0.00244012 4.18212 -0.0159747 4.64173 0.0327013C5.41618 0.130634 5.43492 1.24767 4.67447 1.3862C3.8452 1.53726 2.3482 1.14755 1.72421 1.88378C1.28765 2.39892 1.39058 3.62938 1.3905 4.27952L1.39013 9.66736C1.39013 10.231 1.37846 10.7807 1.42506 11.3449C1.48123 12.0248 1.92446 12.4926 2.60621 12.5693C3.14913 12.6304 3.70966 12.6109 4.25799 12.611L9.61388 12.6113C10.195 12.6112 10.8001 12.6412 11.377 12.568C13.3392 12.319 12.1947 9.70475 12.7558 9.00046C12.8662 8.86191 13.0573 8.79826 13.2285 8.78153C13.4113 8.76365 13.6029 8.81518 13.7453 8.93283C13.9214 9.07839 13.9714 9.25809 13.9936 9.47445C14.0069 9.79167 13.996 10.2067 13.9924 10.5251C13.9815 11.4714 13.9481 12.3379 13.2722 13.0729C12.6072 13.796 11.9487 13.9528 11.0001 13.9858C10.4148 14.0062 9.81089 13.9972 9.22406 13.9974L4.05718 13.997C2.97992 13.9953 1.78126 14.0927 0.933736 13.2992C0.208461 12.6202 0.0333908 12.0302 0.0123049 11.057C-0.00153913 10.4175 0.0004112 9.77421 0.00013432 9.13231L0.000256025 3.99868C0.000809785 3.37545 -0.0133567 2.65819 0.108087 2.05309C0.254347 1.32431 0.913296 0.590609 1.57902 0.285934C2.03923 0.0753118 2.65919 0.0346438 3.16338 0.0172543ZM12.7932 0.00326689C13.2004 0.00324049 13.6457 -0.0582081 13.8866 0.353074C14.0178 0.57714 13.9776 0.925443 13.9818 1.18174C13.9863 1.46096 13.9821 1.74451 13.9823 2.02378L13.9818 5.14249C13.9814 5.48349 14.027 5.9166 13.7631 6.14739C13.6105 6.28074 13.4622 6.358 13.2568 6.35355C13.0783 6.35085 12.9088 6.27504 12.7878 6.14387C12.714 6.06302 12.6145 5.88688 12.6071 5.7763C12.5532 4.96902 12.5934 4.15255 12.5815 3.34273C12.577 3.03285 12.5766 2.70485 12.5915 2.3967C12.2948 2.65347 12.0353 2.94405 11.7522 3.21599C11.1651 3.78003 10.6064 4.39573 10.0128 4.95104C9.85398 5.09269 9.66854 5.30286 9.51287 5.45787L8.34644 6.61615C8.05581 6.90491 7.74948 7.35486 7.3055 7.33632C7.11346 7.3268 6.93344 7.23994 6.8065 7.09562C6.68367 6.95637 6.6226 6.77325 6.63721 6.58818C6.64793 6.45653 6.68742 6.34428 6.77158 6.24263C7.00411 5.96181 7.2862 5.70045 7.54526 5.44194L10.6315 2.35863C10.9079 2.08178 11.3016 1.66368 11.5892 1.41369C10.8561 1.403 10.1229 1.40078 9.38982 1.407C9.06089 1.40742 8.63666 1.42778 8.31565 1.39702C7.49264 1.31814 7.33882 0.26467 8.21622 0.0138487C8.2941 0.00676542 8.42404 0.00428332 8.50466 0.00423992C9.59569 0.00362265 10.6876 0.00507613 11.7785 0.0041183L12.7932 0.00326689Z" fill="${TOOLBAR_ICON_COLOR}"/>
</svg>`

function managerToggleIcon(mode: ManagerKind): string {
  const label = mode === 'manager' ? 'M' : 'IC'
  return `<svg width="20" height="14" viewBox="0 0 20 14" fill="none" xmlns="http://www.w3.org/2000/svg">
  <text x="10" y="11" text-anchor="middle" font-family="Inter, sans-serif" font-size="11" font-weight="400" fill="${TOOLBAR_ICON_COLOR}">${label}</text>
</svg>`
}

const BULK_ICON_DARK = `<svg width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M38.3296 31.8374C38.2405 33.5987 36.7839 34.9993 35.0005 34.9995L26.6665 34.9995L26.4956 34.9956C24.7911 34.9092 23.4236 33.542 23.3374 31.8374L23.3335 31.6665L23.3335 28.3325C23.3339 26.5494 24.7346 25.0927 26.4956 25.0034L26.6665 24.9995L30.0005 24.9995C30.0001 23.1589 28.5072 21.6665 26.6665 21.6665L25.8335 21.6665C23.7109 21.6664 21.8539 20.5325 20.8335 18.8374C19.8131 20.5324 17.9562 21.6665 15.8335 21.6665L15.0005 21.6665C13.1598 21.6666 11.6668 23.1589 11.6665 24.9995L15.0005 24.9995C16.841 24.9997 18.3331 26.492 18.3335 28.3325L18.3335 31.6665L18.3296 31.8374C18.2405 33.5987 16.7839 34.9993 15.0005 34.9995L6.6665 34.9995L6.49561 34.9956C4.79108 34.9092 3.42364 33.542 3.3374 31.8374L3.3335 31.6665L3.3335 28.3325C3.33388 26.5494 4.7346 25.0927 6.49561 25.0034L6.6665 24.9995L10.0005 24.9995C10.0008 22.2384 12.2393 19.9996 15.0005 19.9995L15.8335 19.9995C18.1346 19.9995 20.0003 18.1345 20.0005 15.8335L20.0005 14.9995L16.6665 14.9995L16.4956 14.9956C14.7911 14.9092 13.4236 13.542 13.3374 11.8374L13.3335 11.6665L13.3335 8.33252C13.3339 6.54942 14.7346 5.09266 16.4956 5.00342L16.6665 4.99951L25.0005 4.99951C26.841 4.99969 28.3331 6.49202 28.3335 8.33252L28.3335 11.6665L28.3296 11.8374C28.2405 13.5987 26.7839 14.9993 25.0005 14.9995L21.6665 14.9995L21.6665 15.8335C21.6668 18.1344 23.5326 19.9993 25.8335 19.9995L26.6665 19.9995C29.4277 19.9995 31.6661 22.2384 31.6665 24.9995L35.0005 24.9995C36.841 24.9997 38.3331 26.492 38.3335 28.3325L38.3335 31.6665L38.3296 31.8374ZM16.6665 28.3325C16.6661 27.4125 15.9206 26.6667 15.0005 26.6665L6.6665 26.6665C5.74657 26.6668 5.00089 27.4126 5.00049 28.3325L5.00049 31.6665C5.00066 32.5866 5.74643 33.3322 6.6665 33.3325L15.0005 33.3325C15.9207 33.3323 16.6663 32.5867 16.6665 31.6665L16.6665 28.3325ZM26.6665 8.33252C26.6661 7.41249 25.9206 6.66668 25.0005 6.66651L16.6665 6.6665C15.7466 6.66685 15.0009 7.4126 15.0005 8.33252L15.0005 11.6665C15.0007 12.5866 15.7464 13.3322 16.6665 13.3325L25.0005 13.3325C25.9207 13.3323 26.6663 12.5867 26.6665 11.6665L26.6665 8.33252ZM36.6665 28.3325C36.6661 27.4125 35.9206 26.6667 35.0005 26.6665L26.6665 26.6665C25.7466 26.6668 25.0009 27.4126 25.0005 28.3325L25.0005 31.6665C25.0007 32.5866 25.7464 33.3322 26.6665 33.3325L35.0005 33.3325C35.9207 33.3323 36.6663 32.5867 36.6665 31.6665L36.6665 28.3325Z" fill="black" fill-opacity="0.9"/>
<path d="M35.2 18.2L35.5149 16.9403C35.6412 16.4353 36.3588 16.4353 36.4851 16.9403L36.8 18.2L38.0597 18.5149C38.5647 18.6412 38.5647 19.3588 38.0597 19.4851L36.8 19.8L36.4851 21.0597C36.3588 21.5647 35.6412 21.5647 35.5149 21.0597L35.2 19.8L33.9403 19.4851C33.4353 19.3588 33.4353 18.6412 33.9403 18.5149L35.2 18.2Z" fill="black" fill-opacity="0.9"/>
<path d="M7 13L9.05972 13.5149C9.56469 13.6412 9.56469 14.3588 9.05972 14.4851L7 15L6.48507 17.0597C6.35883 17.5647 5.64117 17.5647 5.51493 17.0597L5 15L2.94029 14.4851C2.43531 14.3588 2.43531 13.6412 2.94029 13.5149L5 13L5.51493 10.9403C5.64117 10.4353 6.35883 10.4353 6.48507 10.9403L7 13Z" fill="black" fill-opacity="0.9"/>
</svg>`

const AVATAR_UPLOAD_UI_HTML = `
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      html, body {
        margin: 0;
        padding: 0;
        width: 100%;
        height: 100%;
        background: #ffffff;
        font-family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      }
      .root {
        box-sizing: border-box;
        width: 100%;
        height: 100%;
        padding: 16px;
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      .title {
        color: #1f2937;
        font-size: 13px;
        line-height: 1.35;
      }
      .actions { display: flex; gap: 8px; }
      button {
        appearance: none;
        border: none;
        border-radius: 8px;
        padding: 9px 12px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
      }
      #choose {
        background: #8d46ff;
        color: #ffffff;
      }
      #secondary {
        background: #f3f4f6;
        color: #4b5563;
      }
      .hint {
        color: #6b7280;
        font-size: 11px;
      }
    </style>
  </head>
  <body>
    <div class="root">
      <div class="title">Upload a profile image to use on this people chip.</div>
      <div class="actions">
        <button id="choose" type="button">Choose image</button>
        <button id="secondary" type="button" style="display:none"></button>
      </div>
      <div class="hint">Supported: PNG, JPG, WEBP, GIF. Image is cropped to a square avatar.</div>
      <input id="file" type="file" accept="image/*" style="display:none" />
    </div>
    <script>
      const MAX_EDGE = 128;
      const TARGET_MAX_DATA_URI_LENGTH = 180000;
      const post = (payload) => parent.postMessage({ pluginMessage: payload }, '*');
      const fileInput = document.getElementById('file');
      const chooseButton = document.getElementById('choose');
      const secondaryButton = document.getElementById('secondary');
      let hasImage = false;

      async function fileToDataUri(file) {
        return await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onerror = () => reject(new Error('Failed to read file.'));
          reader.onload = () => {
            if (typeof reader.result === 'string') resolve(reader.result);
            else reject(new Error('Unexpected file result.'));
          };
          reader.readAsDataURL(file);
        });
      }

      async function loadImage(src) {
        return await new Promise((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = () => reject(new Error('Could not decode image.'));
          image.src = src;
        });
      }

      async function normalizeAvatar(file) {
        const sourceDataUri = await fileToDataUri(file);
        const image = await loadImage(sourceDataUri);
        const side = Math.min(image.width, image.height);
        const sx = (image.width - side) / 2;
        const sy = (image.height - side) / 2;

        const canvas = document.createElement('canvas');
        canvas.width = MAX_EDGE;
        canvas.height = MAX_EDGE;
        const context = canvas.getContext('2d');
        if (!context) throw new Error('Could not prepare image canvas.');

        context.imageSmoothingEnabled = true;
        context.imageSmoothingQuality = 'high';
        context.drawImage(image, sx, sy, side, side, 0, 0, MAX_EDGE, MAX_EDGE);

        let quality = 0.9;
        let output = canvas.toDataURL('image/jpeg', quality);
        while (output.length > TARGET_MAX_DATA_URI_LENGTH && quality >= 0.45) {
          quality -= 0.1;
          output = canvas.toDataURL('image/jpeg', quality);
        }
        return output;
      }

      function syncSecondaryButton() {
        if (hasImage) {
          secondaryButton.style.display = 'inline-flex';
          secondaryButton.textContent = 'Remove image';
        } else {
          secondaryButton.style.display = 'none';
          secondaryButton.textContent = '';
        }
      }

      chooseButton.addEventListener('click', () => fileInput.click());
      secondaryButton.addEventListener('click', () => {
        if (!hasImage) return;
        post({ type: 'avatar-remove-image' });
      });

      window.onmessage = (event) => {
        const pluginMessage = event.data && event.data.pluginMessage;
        if (!pluginMessage || pluginMessage.type !== 'avatar-upload-init') return;
        hasImage = Boolean(pluginMessage.hasImage);
        syncSecondaryButton();
      };

      fileInput.addEventListener('change', async () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;
        try {
          const dataUri = await normalizeAvatar(file);
          post({ type: 'avatar-uploaded', dataUri });
        } catch (error) {
          const message = error instanceof Error ? error.message : 'Could not process this image.';
          post({ type: 'avatar-upload-error', message });
        }
      });
    </script>
  </body>
</html>
`

const LINK_EDITOR_UI_HTML = `
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      html, body {
        margin: 0;
        padding: 0;
        width: 100%;
        height: 100%;
        background: #ffffff;
        font-family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      }
      .root {
        box-sizing: border-box;
        width: 100%;
        height: 100%;
        padding: 14px;
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      input {
        width: 100%;
        box-sizing: border-box;
        border: 1px solid #d1d5db;
        border-radius: 8px;
        padding: 9px 10px;
        font-size: 12px;
        color: #111827;
      }
      .actions { display: flex; gap: 8px; }
      button {
        appearance: none;
        border: none;
        border-radius: 8px;
        padding: 9px 12px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
      }
      #save {
        background: #8d46ff;
        color: #ffffff;
      }
      #cancel {
        background: #f3f4f6;
        color: #4b5563;
      }
    </style>
  </head>
  <body>
    <div class="root">
      <input id="url" type="text" placeholder="https://example.com" />
      <div class="actions">
        <button id="save" type="button">Save link</button>
        <button id="cancel" type="button">Cancel</button>
      </div>
    </div>
    <script>
      const post = (payload) => parent.postMessage({ pluginMessage: payload }, '*');
      const input = document.getElementById('url');
      const saveButton = document.getElementById('save');
      const cancelButton = document.getElementById('cancel');

      saveButton.addEventListener('click', () => {
        post({ type: 'link-save', url: input.value || '' });
      });
      cancelButton.addEventListener('click', () => {
        post({ type: 'link-cancel' });
      });
      input.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          post({ type: 'link-save', url: input.value || '' });
        }
      });
      window.onmessage = (event) => {
        const pluginMessage = event.data && event.data.pluginMessage;
        if (!pluginMessage || pluginMessage.type !== 'link-editor-init') return;
        input.value = pluginMessage.url || '';
        input.focus();
        input.setSelectionRange(input.value.length, input.value.length);
      };
    </script>
  </body>
</html>
`

function roleColor(role: string): string {
  const normalized = role.trim().toLowerCase()
  if (normalized.includes('design')) return '#8F5BFF'
  if (normalized.includes('tpm')) return '#1F9D73'
  if (normalized.includes('pm') || normalized.includes('product')) return '#2A9D8F'
  if (normalized.includes('engineer') || normalized.includes('eng') || normalized.includes('dev')) return '#111111'
  if (normalized.includes('ops') || normalized.includes('operations')) return '#5C6BC0'
  if (normalized.includes('brand')) return '#3B82F6'
  if (normalized.includes('writing') || normalized.includes('writer')) return '#E76F51'
  if (normalized.includes('research') || normalized.includes('uxr')) return '#7C6AE6'
  if (normalized.includes('data') || normalized.includes('analyst')) return '#3FA7A3'
  if (normalized.length === 0) return '#A1A1AA'
  return '#52525B'
}

function initialsFromName(name: string): string {
  const parts = name
    .trim()
    .split(/\s+/)
    .filter(Boolean)

  if (parts.length === 0) return '?'
  if (parts.length === 1) return parts[0].slice(0, 1).toUpperCase()
  return `${parts[0][0]}${parts[1][0]}`.toUpperCase()
}

function roleFieldWidth(role: string): number {
  const text = role.trim()
  const minWidth = 20
  const maxWidth = 120
  const estimatedCharWidth = 4.8
  // Input has intrinsic inset/selection affordances, so keep a small fixed buffer.
  const uiBuffer = 6
  const estimated = Math.ceil(Math.max(text.length, 1) * estimatedCharWidth + uiBuffer)
  return Math.max(minWidth, Math.min(maxWidth, estimated))
}

function textFieldWidth(
  value: string,
  placeholder: string,
  minWidth: number,
  maxWidth: number,
  estimatedCharWidth: number,
): number {
  const text = value.trim().length > 0 ? value.trim() : placeholder
  const uiBuffer = 8
  const estimated = Math.ceil(Math.max(text.length, 1) * estimatedCharWidth + uiBuffer)
  return Math.max(minWidth, Math.min(maxWidth, estimated))
}

function fittedFontSizeForWidth(
  value: string,
  placeholder: string,
  maxWidth: number,
  baseSize: number,
  minSize: number,
): number {
  const text = value.trim().length > 0 ? value.trim() : placeholder
  if (text.length <= 1) return baseSize

  // Empirical approximation for Inter width in widget inputs.
  const charFactor = 0.56
  const safetyPadding = 10
  const estimatedWidthAtBase = text.length * baseSize * charFactor + safetyPadding
  if (estimatedWidthAtBase <= maxWidth) return baseSize

  const scaled = Math.floor((maxWidth - safetyPadding) / (text.length * charFactor))
  return Math.max(minSize, Math.min(baseSize, scaled))
}

function asAvatarUploadMessage(message: unknown): AvatarUploadMessage | null {
  if (!message || typeof message !== 'object') return null
  const record = message as Record<string, unknown>
  const type = record.type
  if (type === 'avatar-remove-image') {
    return { type: 'avatar-remove-image' }
  }
  if (type === 'avatar-upload-error') {
    const text = typeof record.message === 'string' ? record.message : 'Could not upload image.'
    return { type: 'avatar-upload-error', message: text }
  }
  if (type === 'avatar-uploaded' && typeof record.dataUri === 'string') {
    return { type: 'avatar-uploaded', dataUri: record.dataUri }
  }
  return null
}

function asLinkEditMessage(message: unknown): LinkEditMessage | null {
  if (!message || typeof message !== 'object') return null
  const record = message as Record<string, unknown>
  const type = record.type
  if (type === 'link-cancel') return { type: 'link-cancel' }
  if (type === 'link-save') {
    return { type: 'link-save', url: typeof record.url === 'string' ? record.url : '' }
  }
  return null
}

function normalizeExternalUrl(value: string): string {
  const trimmed = value.trim()
  if (!trimmed) return ''
  if (/^(https?:\/\/|mailto:|tel:)/i.test(trimmed)) return trimmed
  return `https://${trimmed}`
}

function getCenterPoint(node: SceneNode): Point {
  const absoluteX = node.absoluteTransform[0][2]
  const absoluteY = node.absoluteTransform[1][2]
  return {
    x: absoluteX + node.width / 2,
    y: absoluteY + node.height / 2,
  }
}

function getTopLeftPoint(node: SceneNode): Point {
  return {
    x: node.absoluteTransform[0][2],
    y: node.absoluteTransform[1][2],
  }
}

function getParentTopLeft(parent: BaseNode & ChildrenMixin): Point {
  if ('absoluteTransform' in parent) {
    return {
      x: parent.absoluteTransform[0][2],
      y: parent.absoluteTransform[1][2],
    }
  }
  return { x: 0, y: 0 }
}

function containsPoint(section: SectionNode, point: Point): boolean {
  const sectionX = section.absoluteTransform[0][2]
  const sectionY = section.absoluteTransform[1][2]
  return (
    point.x >= sectionX &&
    point.x <= sectionX + section.width &&
    point.y >= sectionY &&
    point.y <= sectionY + section.height
  )
}

function findContainingSectionByGeometry(node: SceneNode): SectionNode | null {
  const center = getCenterPoint(node)
  const sections = figma.currentPage.findAllWithCriteria({ types: ['SECTION'] })
  let containing: SectionNode | null = null

  for (const section of sections) {
    if (!containsPoint(section, center)) continue
    if (!containing || section.width * section.height < containing.width * containing.height) {
      containing = section
    }
  }
  return containing
}

function isPeopleChipWidgetNode(node: SceneNode, widgetId: string): node is WidgetNode {
  return (
    node.type === 'WIDGET' &&
    node.widgetId === widgetId &&
    node.widgetSyncedState['widget-kind'] === 'people-chip'
  )
}

function getPeopleChipWidgetsInSection(section: SectionNode, widgetId: string): WidgetNode[] {
  return section.findAll((node) => isPeopleChipWidgetNode(node, widgetId)) as WidgetNode[]
}

function classifyRole(role: string): 'Design' | 'Eng' | 'PM' | 'TPM' | 'Ops' | 'Brand' | 'Writing' | 'Research' | 'Data' | 'Other' {
  const normalized = role.trim().toLowerCase()
  // Explicit canonical labels from the role dropdown.
  if (normalized === 'design') return 'Design'
  if (normalized === 'eng') return 'Eng'
  if (normalized === 'tpm') return 'TPM'
  if (normalized === 'pm') return 'PM'
  if (normalized === 'ops') return 'Ops'
  if (normalized === 'brand') return 'Brand'
  if (normalized === 'writing') return 'Writing'
  if (normalized === 'research') return 'Research'
  if (normalized === 'data') return 'Data'
  // Fallback aliases/typed variants.
  if (normalized.includes('design')) return 'Design'
  if (normalized.includes('tpm')) return 'TPM'
  if (normalized.includes('engineer') || normalized.includes('eng') || normalized.includes('dev')) return 'Eng'
  if (normalized.includes('pm') || normalized.includes('product')) return 'PM'
  if (normalized.includes('ops') || normalized.includes('operations')) return 'Ops'
  if (normalized.includes('brand')) return 'Brand'
  if (normalized.includes('writing') || normalized.includes('writer')) return 'Writing'
  if (normalized.includes('research') || normalized.includes('uxr')) return 'Research'
  if (normalized.includes('data') || normalized.includes('analyst')) return 'Data'
  return 'Other'
}

function buildRoleBreakdown(widgets: WidgetNode[]): string[] {
  const buckets: Record<'Design' | 'Eng' | 'PM' | 'TPM' | 'Ops' | 'Brand' | 'Writing' | 'Research' | 'Data' | 'Other', number> = {
    Design: 0,
    Eng: 0,
    PM: 0,
    TPM: 0,
    Ops: 0,
    Brand: 0,
    Writing: 0,
    Research: 0,
    Data: 0,
    Other: 0,
  }
  for (const widgetNode of widgets) {
    const role = String(widgetNode.widgetSyncedState['role'] ?? '')
    buckets[classifyRole(role)] += 1
  }

  const parts: string[] = []
  if (buckets.Design > 0) parts.push(`${buckets.Design} Design`)
  if (buckets.Eng > 0) parts.push(`${buckets.Eng} Eng`)
  if (buckets.PM > 0) parts.push(`${buckets.PM} PM`)
  if (buckets.TPM > 0) parts.push(`${buckets.TPM} TPM`)
  if (buckets.Ops > 0) parts.push(`${buckets.Ops} Ops`)
  if (buckets.Brand > 0) parts.push(`${buckets.Brand} Brand`)
  if (buckets.Writing > 0) parts.push(`${buckets.Writing} Writing`)
  if (buckets.Research > 0) parts.push(`${buckets.Research} Research`)
  if (buckets.Data > 0) parts.push(`${buckets.Data} Data`)
  if (buckets.Other > 0) parts.push(`${buckets.Other} Other`)
  return parts
}

function roleTitleForCounter(role: 'Design' | 'Eng' | 'PM' | 'TPM' | 'Ops' | 'Brand' | 'Writing' | 'Research' | 'Data' | 'Other', count: number): string {
  if (role === 'Eng') return `${count} ${count === 1 ? 'Engineer' : 'Engineers'}`
  if (role === 'Other') return `${count} Other`
  return `${count} ${role}`
}

function buildCounterRoleDetails(widgets: WidgetNode[]): CounterRoleDetail[] {
  const roleOrder: Array<'Design' | 'PM' | 'Eng' | 'TPM' | 'Ops' | 'Brand' | 'Writing' | 'Research' | 'Data' | 'Other'> = [
    'Design',
    'PM',
    'Eng',
    'TPM',
    'Ops',
    'Brand',
    'Writing',
    'Research',
    'Data',
    'Other',
  ]

  const buckets: Record<
    'Design' | 'Eng' | 'PM' | 'TPM' | 'Ops' | 'Brand' | 'Writing' | 'Research' | 'Data' | 'Other',
    { count: number; active: number; open: number; newCount: number }
  > = {
    Design: { count: 0, active: 0, open: 0, newCount: 0 },
    Eng: { count: 0, active: 0, open: 0, newCount: 0 },
    PM: { count: 0, active: 0, open: 0, newCount: 0 },
    TPM: { count: 0, active: 0, open: 0, newCount: 0 },
    Ops: { count: 0, active: 0, open: 0, newCount: 0 },
    Brand: { count: 0, active: 0, open: 0, newCount: 0 },
    Writing: { count: 0, active: 0, open: 0, newCount: 0 },
    Research: { count: 0, active: 0, open: 0, newCount: 0 },
    Data: { count: 0, active: 0, open: 0, newCount: 0 },
    Other: { count: 0, active: 0, open: 0, newCount: 0 },
  }

  for (const widgetNode of widgets) {
    const role = classifyRole(String(widgetNode.widgetSyncedState['role'] ?? ''))
    const statusValue = normalizeStatus(widgetNode.widgetSyncedState['status'])
    const bucket = buckets[role]
    bucket.count += 1
    if (statusValue === 'active') bucket.active += 1
    else if (statusValue === 'open-headcount') bucket.open += 1
    else bucket.newCount += 1
  }

  return roleOrder
    .filter((role) => buckets[role].count > 0)
    .sort((a, b) => {
      const countDelta = buckets[b].count - buckets[a].count
      if (countDelta !== 0) return countDelta
      return roleOrder.indexOf(a) - roleOrder.indexOf(b)
    })
    .map((role) => {
      const bucket = buckets[role]
      const parts: string[] = []
      if (bucket.active > 0) parts.push(`${bucket.active} Active`)
      if (bucket.open > 0) parts.push(`${bucket.open} Open`)
      if (bucket.newCount > 0) parts.push(`${bucket.newCount} New`)
      return {
        title: roleTitleForCounter(role, bucket.count),
        color: roleColor(role === 'Eng' ? 'Eng' : role),
        statusSummary: parts.join(', '),
      } satisfies CounterRoleDetail
    })
}

function buildCounterLocationDetails(widgets: WidgetNode[]): CounterRoleDetail[] {
  const buckets = new Map<string, number>()
  for (const widgetNode of widgets) {
    const rawLocation = String(widgetNode.widgetSyncedState['location'] ?? '').trim()
    if (rawLocation.length === 0) continue
    if (/^Q[1-4]$/i.test(rawLocation)) continue
    const location = rawLocation
    buckets.set(location, (buckets.get(location) ?? 0) + 1)
  }

  return Array.from(buckets.entries())
    .sort((a, b) => {
      if (b[1] !== a[1]) return b[1] - a[1]
      return a[0].localeCompare(b[0])
    })
    .map(([location, count]) => ({
      title: `${count} ${location}`,
      color: '#4B5563',
      statusSummary: '',
    }))
}

function matchesPreset(widgetNode: WidgetNode, preset: CounterPreset): boolean {
  if (preset === 'all' || preset === 'locations') return true
  const chipStatus = widgetNode.widgetSyncedState['status']
  return chipStatus === preset
}

function matchesPeopleFilter(widgetNode: WidgetNode, peopleFilter: CounterPeopleFilter): boolean {
  if (peopleFilter === 'all') return true
  const managerKind = String(widgetNode.widgetSyncedState['manager-kind'] ?? '')
    .trim()
    .toLowerCase()
  if (peopleFilter === 'managers') return managerKind === 'manager' || managerKind === 'm'
  return managerKind === '' || managerKind === 'ic'
}

function normalizeStatus(value: unknown): Status {
  const normalized = String(value ?? '')
    .trim()
    .toLowerCase()
  if (normalized.includes('open')) return 'open-headcount'
  if (normalized.includes('new')) return 'new-headcount'
  return 'active'
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null
  return value as Record<string, unknown>
}

function normalizeKey(key: string): string {
  return key.toLowerCase().replace(/[\s_-]/g, '')
}

function profileFromObject(value: unknown): ChipProfile | null {
  const record = asRecord(value)
  if (!record) return null

  const byKey: Record<string, unknown> = {}
  for (const key in record) {
    if (Object.prototype.hasOwnProperty.call(record, key)) {
      byKey[normalizeKey(key)] = record[key]
    }
  }

  const name = String(byKey['name'] ?? byKey['fullname'] ?? byKey['person'] ?? '').trim()
  if (!name) return null

  return {
    name,
    projectArea: String(
      byKey['projectarea'] ??
      byKey['team'] ??
      byKey['teamarea'] ??
      byKey['org'] ??
      byKey['department'] ??
      '',
    ).trim(),
    role: String(byKey['role'] ?? byKey['title'] ?? byKey['function'] ?? 'Eng').trim() || 'Eng',
    location: String(byKey['location'] ?? byKey['city'] ?? byKey['office'] ?? '').trim(),
    status: normalizeStatus(byKey['status'] ?? byKey['state']),
  }
}

function parseJsonProfiles(input: string): ChipProfile[] {
  const importResult = parseJsonImport(input)
  if (importResult && importResult.cards.length > 0) {
    return importResult.cards.map((card) => card.profile)
  }
  return []
}

function layoutIndexedImportNodes(
  nodes: IndexedImportNode[],
  connections: ParsedConnection[],
): ImportParseResult {
  const adjacency = new Map<string, string[]>()
  const indegree = new Map<string, number>()
  const levelById = new Map<string, number>()
  const orderById = new Map<string, number>()

  for (const node of nodes) {
    adjacency.set(node.id, [])
    indegree.set(node.id, 0)
    orderById.set(node.id, node.order)
  }

  for (const connection of connections) {
    if (!adjacency.has(connection.fromSourceId) || !adjacency.has(connection.toSourceId)) continue
    adjacency.get(connection.fromSourceId)?.push(connection.toSourceId)
    indegree.set(connection.toSourceId, (indegree.get(connection.toSourceId) ?? 0) + 1)
  }

  const roots = nodes
    .filter((node) => (indegree.get(node.id) ?? 0) === 0)
    .sort((a, b) => a.order - b.order)

  const queue = [...roots.map((node) => node.id)]
  for (const rootId of queue) levelById.set(rootId, 0)

  while (queue.length > 0) {
    const currentId = queue.shift()
    if (!currentId) continue
    const currentLevel = levelById.get(currentId) ?? 0
    const children = adjacency.get(currentId) ?? []
    for (const childId of children) {
      const nextLevel = currentLevel + 1
      const existingLevel = levelById.get(childId)
      if (existingLevel === undefined || nextLevel < existingLevel) {
        levelById.set(childId, nextLevel)
      }
      queue.push(childId)
    }
  }

  for (const node of nodes) {
    if (!levelById.has(node.id)) levelById.set(node.id, 0)
  }

  const levels = new Map<number, IndexedImportNode[]>()
  let maxColumns = 0
  for (const node of nodes) {
    const level = levelById.get(node.id) ?? 0
    const list = levels.get(level) ?? []
    list.push(node)
    list.sort((a, b) => (orderById.get(a.id) ?? 0) - (orderById.get(b.id) ?? 0))
    levels.set(level, list)
    if (list.length > maxColumns) maxColumns = list.length
  }

  const colSpacing = 320
  const rowSpacing = 150
  const cards: ParsedCardCandidate[] = []

  for (const [level, levelNodes] of levels.entries()) {
    const offset = ((maxColumns - levelNodes.length) * colSpacing) / 2
    for (let i = 0; i < levelNodes.length; i += 1) {
      const node = levelNodes[i]
      cards.push({
        sourceNodeId: node.id,
        x: offset + i * colSpacing,
        y: level * rowSpacing,
        width: PEOPLE_CHIP_WIDTH,
        height: 80,
        profile: node.profile,
      })
    }
  }

  return {
    cards,
    connections,
    containsRasterImage: false,
  }
}

function parseJsonImport(input: string): ImportParseResult | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(input)
  } catch {
    return null
  }

  if (Array.isArray(parsed) && parsed.every((row) => Array.isArray(row))) {
    const rows = parsed as unknown[][]
    const nodes: IndexedImportNode[] = []
    const parentById = new Map<string, string>()
    const idByName = new Map<string, string>()

    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i]
      const name = String(row[0] ?? '').trim()
      if (!name) continue
      const parentName = String(row[1] ?? '').trim()
      const role = String(row[2] ?? '').trim() || 'Eng'
      const id = `arr-${i}-${name}`
      nodes.push({
        id,
        order: i,
        profile: {
          name,
          projectArea: '',
          role,
          location: '',
          status: 'active',
        },
      })
      if (!idByName.has(name)) idByName.set(name, id)
      parentById.set(id, parentName)
    }

    const connections: ParsedConnection[] = []
    for (const node of nodes) {
      const parentName = parentById.get(node.id)
      if (!parentName) continue
      const parentId = idByName.get(parentName)
      if (!parentId || parentId === node.id) continue
      connections.push({ fromSourceId: parentId, toSourceId: node.id })
    }

    if (nodes.length > 0) return layoutIndexedImportNodes(nodes, connections)
  }

  const rootRecord = asRecord(parsed)
  const treeRoot =
    rootRecord && !Array.isArray(rootRecord['children'])
      ? null
      : asRecord(parsed)
  if (treeRoot && Array.isArray(treeRoot['children'])) {
    const nodes: IndexedImportNode[] = []
    const connections: ParsedConnection[] = []
    let order = 0

    const walk = (value: unknown, parentId: string | null) => {
      const record = asRecord(value)
      if (!record) return
      const profile = profileFromObject(record)
      if (!profile) return
      const rawId = String(record['id'] ?? '').trim()
      const id = rawId || `tree-${order}-${profile.name}`
      nodes.push({ id, profile, order })
      order += 1
      if (parentId) connections.push({ fromSourceId: parentId, toSourceId: id })
      const children = Array.isArray(record['children']) ? (record['children'] as unknown[]) : []
      for (const child of children) walk(child, id)
    }

    walk(parsed, null)
    if (nodes.length > 0) return layoutIndexedImportNodes(nodes, connections)
  }

  const parsedRecord = asRecord(parsed)
  const rootArray =
    Array.isArray(parsed)
      ? parsed
      : Array.isArray(parsedRecord?.people)
        ? parsedRecord.people
        : Array.isArray(parsedRecord?.chips)
          ? parsedRecord.chips
          : Array.isArray(parsedRecord?.data)
            ? parsedRecord.data
            : []
  const flatNodes: IndexedImportNode[] = rootArray
    .map((item, index) => {
      const profile = profileFromObject(item)
      if (!profile) return null
      return {
        id: `flat-${index}-${profile.name}`,
        order: index,
        profile,
      } satisfies IndexedImportNode
    })
    .filter((node): node is IndexedImportNode => Boolean(node))

  if (flatNodes.length === 0) return null
  return layoutIndexedImportNodes(flatNodes, [])
}

function splitColumns(line: string): string[] {
  if (line.includes('\t')) return line.split('\t').map((v) => v.trim())
  if (line.includes('|')) return line.split('|').map((v) => v.trim())
  return line.split(',').map((v) => v.trim())
}

function parseDelimitedProfiles(input: string): ChipProfile[] {
  const lines = input
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)

  if (lines.length === 0) return []

  const headerCols = splitColumns(lines[0]).map((c) => normalizeKey(c))
  const hasHeader = headerCols.some((c) => ['name', 'role', 'location', 'projectarea', 'team'].includes(c))
  const dataLines = hasHeader ? lines.slice(1) : lines

  const hasExplicitDelimiter = lines.some((line) => line.includes('\t') || line.includes('|') || line.includes(','))
  const parsedRows = hasExplicitDelimiter || hasHeader
    ? dataLines
    .map((line) => {
      const cols = splitColumns(line)
      if (cols.length === 0 || !cols[0]) return null

      if (hasHeader) {
        const map: Record<string, string> = {}
        for (let i = 0; i < headerCols.length; i += 1) map[headerCols[i]] = cols[i] ?? ''
        return {
          name: (map['name'] ?? map['fullname'] ?? '').trim(),
          projectArea: (map['projectarea'] ?? map['team'] ?? map['teamarea'] ?? '').trim(),
          role: (map['role'] ?? map['title'] ?? 'Eng').trim() || 'Eng',
          location: (map['location'] ?? map['city'] ?? '').trim(),
          status: normalizeStatus(map['status'] ?? map['state']),
        } satisfies ChipProfile
      }

      return {
        name: cols[0] ?? '',
        projectArea: cols[1] ?? '',
        role: cols[2] || 'Eng',
        location: cols[3] ?? '',
        status: normalizeStatus(cols[4]),
      } satisfies ChipProfile
    })
    .filter((item): item is ChipProfile => Boolean(item && item.name))
    : []

  if (parsedRows.length > 0) return parsedRows

  const plainLines = lines
    .map((line) => line.trim())
    .filter((line) => Boolean(line) && !line.includes(',') && !line.includes('|') && !line.includes('\t'))

  const isRoleLike = (value: string): boolean => {
    const normalized = value.toLowerCase().trim()
    return (
      normalized === 'design' ||
      normalized === 'designer' ||
      normalized === 'eng' ||
      normalized === 'engineer' ||
      normalized === 'pm' ||
      normalized === 'ops' ||
      normalized === 'operations' ||
      normalized === 'writing' ||
      normalized === 'research' ||
      normalized === 'data' ||
      normalized.includes('design') ||
      normalized.includes('engineer') ||
      normalized.includes('eng') ||
      normalized.includes('pm') ||
      normalized.includes('ops') ||
      normalized.includes('operations') ||
      normalized.includes('writing') ||
      normalized.includes('research') ||
      normalized.includes('data')
    )
  }

  const looksLikeName = (value: string): boolean => {
    const words = value.trim().split(/\s+/).filter(Boolean)
    if (words.length === 0 || words.length > 3) return false
    if (isRoleLike(value)) return false
    if (value.toLowerCase().includes('team')) return false
    if (/\d/.test(value)) return false
    return words.every((word) => {
      const first = word.charAt(0)
      const rest = word.slice(1)
      const isInitialUpper = first.toUpperCase() === first && first.toLowerCase() !== first
      const hasMostlyUpperRest = rest === rest.toUpperCase()
      const hasMostlyLowerRest = rest === rest.toLowerCase()
      return isInitialUpper && (hasMostlyLowerRest || hasMostlyUpperRest || word.includes("'"))
    })
  }

  // Handles flattened text copied from many cards (e.g. frames):
  // name blocks plus role-centered triplets (description, role, location).
  const grouped: ChipProfile[] = []
  const roleTriplets: Array<{ description: string; role: string; location: string }> = []
  for (let i = 1; i < plainLines.length - 1; i += 1) {
    if (!isRoleLike(plainLines[i])) continue
    roleTriplets.push({
      description: plainLines[i - 1],
      role: plainLines[i],
      location: plainLines[i + 1],
    })
  }

  const namePool = plainLines.filter((line) => looksLikeName(line))
  const uniqueNames: string[] = []
  for (let i = 0; i < namePool.length; i += 1) {
    if (uniqueNames.includes(namePool[i])) continue
    uniqueNames.push(namePool[i])
  }

  const hasCardLikeSignal =
    roleTriplets.length >= 2 &&
    uniqueNames.length >= roleTriplets.length &&
    roleTriplets.some((triplet) => triplet.location.trim().length > 0)

  if (hasCardLikeSignal) {
    const limit = Math.min(uniqueNames.length, roleTriplets.length)
    for (let i = 0; i < limit; i += 1) {
      grouped.push({
        name: uniqueNames[i],
        projectArea: roleTriplets[i].description ?? '',
        role: roleTriplets[i].role || 'Eng',
        location: roleTriplets[i].location ?? '',
        status: 'active',
      })
    }
  }

  const seen = new Set<string>()
  const deduped = grouped.filter((profile) => {
    const key = `${profile.name}|${profile.projectArea}|${profile.role}|${profile.location}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
  if (deduped.length > 0) return deduped

  // Plain fallback: one chip per meaningful line (best for simple pasted lists).
  const plainFallback = plainLines
    .filter((line) => line.length > 0)
    .filter((line) => !isRoleLike(line))
    .map((line) => ({
      name: line,
      projectArea: '',
      role: 'Eng',
      location: '',
      status: 'active',
    } satisfies ChipProfile))

  const fallbackSeen = new Set<string>()
  return plainFallback.filter((profile) => {
    if (fallbackSeen.has(profile.name)) return false
    fallbackSeen.add(profile.name)
    return true
  })
}

function parseInputProfiles(input: string): ChipProfile[] {
  const trimmed = input.trim()
  if (!trimmed) return []
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      const jsonProfiles = parseJsonProfiles(trimmed)
      if (jsonProfiles.length > 0) return jsonProfiles
    } catch {
      // Fall back to delimited parsing.
    }
  }
  return parseDelimitedProfiles(trimmed)
}

const OCR_TITLE_TOKENS = new Set([
  'ceo',
  'cfo',
  'coo',
  'cio',
  'cto',
  'cmo',
  'vp',
  'svp',
  'evp',
  'president',
  'chief',
  'officer',
  'director',
  'manager',
  'head',
  'lead',
  'marketing',
  'operations',
  'information',
  'security',
  'finance',
  'engineering',
  'technology',
  'product',
  'hr',
  'people',
])

function isLikelyNameToken(token: string): boolean {
  const cleaned = token.replace(/[^A-Za-z'’.-]/g, '')
  if (!cleaned) return false
  if (cleaned.length < 2 || cleaned.length > 18) return false
  const first = cleaned.charAt(0)
  if (!(first.toUpperCase() === first && first.toLowerCase() !== first)) return false
  const lowered = cleaned.toLowerCase()
  if (OCR_TITLE_TOKENS.has(lowered)) return false
  return true
}

function parseOcrProfileFromLine(text: string): { name: string; title: string } | null {
  const cleanedLine = text
    .replace(/\s+/g, ' ')
    .replace(/^[\-*•]\s*/, '')
    .trim()
  if (!cleanedLine) return null

  const commaIndex = cleanedLine.indexOf(',')
  if (commaIndex >= 0) {
    const namePart = cleanedLine.slice(0, commaIndex).trim()
    const titlePart = cleanedLine.slice(commaIndex + 1).trim()
    if (!namePart) return null
    const words = namePart.split(/\s+/).filter(Boolean)
    if (words.length >= 1 && words.length <= 4 && words.every((word) => isLikelyNameToken(word))) {
      return { name: namePart, title: titlePart }
    }
  }

  const tokens = cleanedLine.split(/\s+/).filter(Boolean)
  if (tokens.length === 1 && isLikelyNameToken(tokens[0])) {
    return { name: tokens[0], title: '' }
  }
  if (tokens.length >= 2 && tokens.length <= 4 && tokens.every((token) => isLikelyNameToken(token))) {
    return { name: tokens.join(' '), title: '' }
  }

  if (tokens.length >= 2) {
    const firstLower = tokens[0].toLowerCase()
    const last = tokens[tokens.length - 1]
    if (OCR_TITLE_TOKENS.has(firstLower) && isLikelyNameToken(last)) {
      return { name: last, title: tokens.slice(0, tokens.length - 1).join(' ') }
    }
    const lastLower = tokens[tokens.length - 1].toLowerCase()
    if (OCR_TITLE_TOKENS.has(lastLower) && isLikelyNameToken(tokens[0])) {
      return { name: tokens[0], title: tokens.slice(1).join(' ') }
    }
  }

  return null
}

function parseNameAndTitleFromLine(rawLine: string): { name: string; title: string } | null {
  const trimmed = rawLine.trim()
  if (!trimmed) return null
  const noBullet = trimmed.replace(/^((?:-{1,}|[*•]+)|\d+[.)])\s*/, '').trim()
  if (!noBullet) return null

  const commaIndex = noBullet.indexOf(',')
  if (commaIndex >= 0) {
    const name = noBullet.slice(0, commaIndex).trim()
    const title = noBullet.slice(commaIndex + 1).trim()
    if (!name) return null
    return { name, title }
  }

  return { name: noBullet, title: '' }
}

function parseBulletedHierarchyImport(input: string): ImportParseResult | null {
  const rows = input
    .split('\n')
    .map((line) => line.replace(/\r/g, ''))
    .filter((line) => line.trim().length > 0)

  if (rows.length < 2) return null
  const bulletRegex = /^(\s*)((?:-{1,}|[*•]+)|\d+[.)])\s*(.+)$/
  const hasBullet = rows.some((row) => bulletRegex.test(row))
  if (!hasBullet) return null

  const nodes: IndexedImportNode[] = []
  const connections: ParsedConnection[] = []
  const stack: Array<{ level: number; id: string }> = []
  const seenNames = new Map<string, number>()

  const ensureUniqueId = (name: string, index: number) => {
    const existing = seenNames.get(name) ?? 0
    seenNames.set(name, existing + 1)
    return `hier-${index}-${name}-${existing}`
  }

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i]
    const bulletMatch = row.match(bulletRegex)
    const leadingSpaces = (row.match(/^\s*/)?.[0].length ?? 0)
    const indentLevel = Math.floor(leadingSpaces / 2)
    let level = indentLevel
    if (bulletMatch) {
      const marker = bulletMatch[2]
      if (/^-+$/.test(marker)) {
        // Support markdown-ish outline syntax: "-", "--", "---" = depth 1,2,3...
        level = Math.max(level, marker.length)
      } else {
        level = Math.max(level, 1)
      }
    }

    const parsed = parseNameAndTitleFromLine(row)
    if (!parsed || !parsed.name) continue
    if (nodes.length === 0) level = 0

    const id = ensureUniqueId(parsed.name, i)
    nodes.push({
      id,
      order: i,
      profile: {
        name: parsed.name,
        projectArea: parsed.title,
        role: 'Eng',
        location: '',
        status: 'active',
      },
    })

    while (stack.length > 0 && stack[stack.length - 1].level >= level) {
      stack.pop()
    }
    const parent = stack[stack.length - 1]
    if (parent) {
      connections.push({ fromSourceId: parent.id, toSourceId: id })
    }
    stack.push({ level, id })
  }

  if (nodes.length === 0) return null
  return layoutIndexedImportNodes(nodes, connections)
}

function parseInputImport(input: string): ImportParseResult | null {
  const trimmed = input.trim()
  if (!trimmed) return null
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    const parsedJson = parseJsonImport(trimmed)
    if (parsedJson && parsedJson.cards.length > 0) return parsedJson
  }

  const bulletedHierarchy = parseBulletedHierarchyImport(trimmed)
  if (bulletedHierarchy && bulletedHierarchy.cards.length > 0) {
    return bulletedHierarchy
  }

  const arrowLines = trimmed
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
  const arrowPattern = /^(.+?)\s*(?:→|->)\s*(.+)$/
  const hasArrowSyntax = arrowLines.some((line) => arrowPattern.test(line))
  if (hasArrowSyntax) {
    const namesInOrder: string[] = []
    const idByName = new Map<string, string>()
    const connections: ParsedConnection[] = []
    const connectionKeys = new Set<string>()

    const ensureNode = (nameValue: string) => {
      const cleanName = nameValue.trim()
      if (!cleanName) return null
      if (!idByName.has(cleanName)) {
        const nextId = `graph-${namesInOrder.length}-${cleanName}`
        namesInOrder.push(cleanName)
        idByName.set(cleanName, nextId)
      }
      return idByName.get(cleanName) ?? null
    }

    for (const line of arrowLines) {
      const match = line.match(arrowPattern)
      if (!match) {
        ensureNode(line)
        continue
      }
      const fromName = match[1].trim()
      const toName = match[2].trim()
      const fromId = ensureNode(fromName)
      const toId = ensureNode(toName)
      if (!fromId || !toId || fromId === toId) continue
      const key = `${fromId}->${toId}`
      if (connectionKeys.has(key)) continue
      connectionKeys.add(key)
      connections.push({ fromSourceId: fromId, toSourceId: toId })
    }

    if (namesInOrder.length > 0) {
      const nodes: IndexedImportNode[] = namesInOrder.map((name, index) => ({
        id: idByName.get(name) ?? `graph-${index}-${name}`,
        order: index,
        profile: {
          name,
          projectArea: '',
          role: 'Eng',
          location: '',
          status: 'active',
        },
      }))
      return layoutIndexedImportNodes(nodes, connections)
    }
  }

  const profiles = parseDelimitedProfiles(trimmed)
  if (profiles.length === 0) return null
  return {
    cards: profiles.map((profile, index) => ({
      sourceNodeId: `text-${index}-${profile.name}`,
      x: (index % 2) * 320,
      y: Math.floor(index / 2) * 96,
      width: PEOPLE_CHIP_WIDTH,
      height: 80,
      profile,
    })),
    connections: [],
    containsRasterImage: false,
  }
}

function profileFromSelectionNode(node: SceneNode): ChipProfile | null {
  const texts: string[] = []
  if (node.type === 'TEXT') texts.push(node.characters.trim())
  if ('findAllWithCriteria' in node) {
    const descendants = node.findAllWithCriteria({ types: ['TEXT'] })
    for (const textNode of descendants) {
      const value = textNode.characters.trim()
      if (value) texts.push(value)
    }
  }

  const cleaned = Array.from(new Set(texts.map((t) => t.trim()).filter(Boolean)))
  if (cleaned.length === 0) {
    const fallback = node.name.trim()
    if (!fallback) return null
    return { name: fallback, projectArea: '', role: 'Eng', location: '', status: 'active' }
  }

  return {
    name: cleaned[0] ?? node.name,
    projectArea: cleaned[1] ?? '',
    role: cleaned[2] ?? 'Eng',
    location: cleaned[3] ?? '',
    status: normalizeStatus(cleaned[4]),
  }
}

function parseSelectionProfiles(selection: readonly SceneNode[]): ChipProfile[] {
  return selection
    .map((node) => profileFromSelectionNode(node))
    .filter((item): item is ChipProfile => Boolean(item && item.name))
}

function pluralize(count: number, singular: string, plural: string): string {
  return count === 1 ? singular : plural
}

function detectedSummary(peopleCount: number, connectionCount: number): string {
  const peoplePart = `${peopleCount} ${pluralize(peopleCount, 'person', 'people')}`
  if (connectionCount <= 0) return `${peoplePart} detected`
  const connectionPart = `${connectionCount} ${pluralize(connectionCount, 'connection', 'connections')}`
  return `${peoplePart} & ${connectionPart} detected`
}

function isLikelyOcrPersonName(value: string): boolean {
  return parseOcrProfileFromLine(value) !== null
}

function isLikelyOcrTitle(value: string): boolean {
  const text = value.trim()
  if (!text) return false
  const lowered = text.toLowerCase()
  const titleSignals = [
    'chief',
    'officer',
    'president',
    'vice',
    'head',
    'director',
    'manager',
    'lead',
    'marketing',
    'finance',
    'operations',
    'security',
    'technology',
    'information',
    'product',
    'engineering',
  ]
  return titleSignals.some((term) => lowered.includes(term))
}

function bytesToBase64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  let output = ''

  for (let i = 0; i < bytes.length; i += 3) {
    const first = bytes[i] ?? 0
    const second = bytes[i + 1] ?? 0
    const third = bytes[i + 2] ?? 0
    const combined = (first << 16) | (second << 8) | third

    output += alphabet[(combined >> 18) & 63]
    output += alphabet[(combined >> 12) & 63]
    output += i + 1 < bytes.length ? alphabet[(combined >> 6) & 63] : '='
    output += i + 2 < bytes.length ? alphabet[combined & 63] : '='
  }

  return output
}

function parseOcrOverlayLines(payload: Record<string, unknown>): OcrOverlayLine[] {
  const parsedResults = Array.isArray(payload['ParsedResults'])
    ? (payload['ParsedResults'] as Array<Record<string, unknown>>)
    : []

  const lines: OcrOverlayLine[] = []
  for (const parsed of parsedResults) {
    const textOverlay = parsed['TextOverlay']
    if (!textOverlay || typeof textOverlay !== 'object') continue
    const overlayLines = Array.isArray((textOverlay as Record<string, unknown>)['Lines'])
      ? ((textOverlay as Record<string, unknown>)['Lines'] as Array<Record<string, unknown>>)
      : []

    for (const line of overlayLines) {
      const text = String(line['LineText'] ?? '').trim()
      if (!text) continue
      const words = Array.isArray(line['Words']) ? (line['Words'] as Array<Record<string, unknown>>) : []

      let left = Number(line['MinLeft'] ?? NaN)
      let top = Number(line['MinTop'] ?? NaN)
      let width = Number(line['MaxWidth'] ?? NaN)
      let height = Number(line['MaxHeight'] ?? NaN)

      if (words.length > 0) {
        const numericLeft = words.map((word) => Number(word['Left'] ?? NaN)).filter((value) => Number.isFinite(value))
        const numericTop = words.map((word) => Number(word['Top'] ?? NaN)).filter((value) => Number.isFinite(value))
        const numericWidth = words.map((word) => Number(word['Width'] ?? NaN)).filter((value) => Number.isFinite(value))
        const numericHeight = words.map((word) => Number(word['Height'] ?? NaN)).filter((value) => Number.isFinite(value))
        if (numericLeft.length > 0) left = Math.min(...numericLeft)
        if (numericTop.length > 0) top = Math.min(...numericTop)
        if (numericWidth.length > 0 && numericLeft.length === words.length) {
          const rightEdges = words
            .map((word) => Number(word['Left'] ?? NaN) + Number(word['Width'] ?? NaN))
            .filter((value) => Number.isFinite(value))
          if (rightEdges.length > 0) width = Math.max(...rightEdges) - left
        }
        if (numericHeight.length > 0) height = Math.max(...numericHeight)
      }

      if (!Number.isFinite(left) || !Number.isFinite(top)) continue
      lines.push({
        text,
        left,
        top,
        width: Number.isFinite(width) ? width : Math.max(60, text.length * 8),
        height: Number.isFinite(height) ? height : 20,
      })
    }
  }

  return lines.sort((a, b) => {
    if (Math.abs(a.top - b.top) > 2) return a.top - b.top
    return a.left - b.left
  })
}

function inferImportFromOcrOverlay(lines: OcrOverlayLine[]): ImportParseResult | null {
  const parsedNameLines = lines
    .map((line) => {
      const parsed = parseOcrProfileFromLine(line.text)
      if (!parsed) return null
      return {
        line,
        name: parsed.name.trim(),
        embeddedTitle: parsed.title.trim(),
      }
    })
    .filter(
      (
        value,
      ): value is {
        line: OcrOverlayLine
        name: string
        embeddedTitle: string
      } => Boolean(value && value.name),
    )

  if (parsedNameLines.length === 0) return null

  const dedupedNameLines: Array<{
    line: OcrOverlayLine
    name: string
    embeddedTitle: string
  }> = []
  const seenNames = new Set<string>()
  for (const candidate of parsedNameLines) {
    const normalizedName = candidate.name.toLowerCase()
    if (seenNames.has(normalizedName)) continue
    seenNames.add(normalizedName)
    dedupedNameLines.push(candidate)
  }

  const cards: ParsedCardCandidate[] = dedupedNameLines.map((candidateInfo, index) => {
    const line = candidateInfo.line
    const titleCandidate = lines.find((candidateLine) => {
      if (candidateLine.top <= line.top) return false
      if (candidateLine.top - line.top > 90) return false
      const sameColumn = Math.abs(candidateLine.left - line.left) <= 120
      return sameColumn && !isLikelyOcrPersonName(candidateLine.text) && isLikelyOcrTitle(candidateLine.text)
    })

    return {
      sourceNodeId: `ocr-${index}-${candidateInfo.name}`,
      x: line.left,
      y: line.top,
      width: PEOPLE_CHIP_WIDTH,
      height: 80,
      profile: {
        name: candidateInfo.name,
        projectArea: candidateInfo.embeddedTitle || titleCandidate?.text || '',
        role: 'Eng',
        location: '',
        status: 'active',
      },
    } satisfies ParsedCardCandidate
  })

  const rowThreshold = 70
  const rows: ParsedCardCandidate[][] = []
  for (const card of [...cards].sort((a, b) => a.y - b.y || a.x - b.x)) {
    const row = rows[rows.length - 1]
    if (!row || Math.abs(card.y - row[0].y) > rowThreshold) {
      rows.push([card])
    } else {
      row.push(card)
    }
  }

  const connections: ParsedConnection[] = []
  if (rows.length >= 2) {
    const seenKeys = new Set<string>()
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const parentRow = rows[rowIndex - 1]
      const childRow = rows[rowIndex]
      for (const child of childRow) {
        const parent = [...parentRow].sort((a, b) => Math.abs(a.x - child.x) - Math.abs(b.x - child.x))[0]
        if (!parent) continue
        const key = `${parent.sourceNodeId}->${child.sourceNodeId}`
        if (seenKeys.has(key)) continue
        seenKeys.add(key)
        connections.push({ fromSourceId: parent.sourceNodeId, toSourceId: child.sourceNodeId })
      }
    }
  }

  return {
    cards,
    connections,
    containsRasterImage: true,
  }
}

function pickRasterSelectionNode(selection: readonly SceneNode[]): SceneNode | null {
  let best: SceneNode | null = null
  let bestArea = -1

  const consider = (node: SceneNode) => {
    if (node.type === 'WIDGET' || node.type === 'CONNECTOR') return
    if (!nodeHasRasterImage(node)) return
    const area = node.width * node.height
    if (area > bestArea) {
      best = node
      bestArea = area
    }
  }

  for (const node of selection) {
    consider(node)
    if ('findAll' in node) {
      const descendants = node.findAll((descendant) => {
        if (descendant.type === 'WIDGET' || descendant.type === 'CONNECTOR') return false
        return nodeHasRasterImage(descendant)
      })
      for (const descendant of descendants) consider(descendant)
    }
  }

  return best
}

function ocrParseBody(base64ImageDataUri: string): string {
  return [
    `apikey=${encodeURIComponent(OCR_SPACE_API_KEY)}`,
    `language=${encodeURIComponent('eng')}`,
    `isOverlayRequired=${encodeURIComponent('true')}`,
    `OCREngine=${encodeURIComponent('2')}`,
    `base64Image=${encodeURIComponent(base64ImageDataUri)}`,
  ].join('&')
}

async function parseRasterSelectionWithOcr(selection: readonly SceneNode[]): Promise<ImportParseResult | null> {
  const targetNode = pickRasterSelectionNode(selection)
  if (!targetNode) return null

  const pngBytes = await targetNode.exportAsync({
    format: 'PNG',
    constraint: { type: 'WIDTH', value: 1600 },
  })
  const base64ImageDataUri = `data:image/png;base64,${bytesToBase64(pngBytes)}`
  const fetchFn = (globalThis as { fetch?: (input: string, init?: unknown) => Promise<unknown> }).fetch
  if (!fetchFn) return null

  const response = (await Promise.race([
    fetchFn(OCR_SPACE_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: ocrParseBody(base64ImageDataUri),
    }),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error('OCR request timed out.')), 15000),
    ),
  ])) as {
    ok: boolean
    status: number
    json: () => Promise<unknown>
  }

  if (!response.ok) return null
  const payload = (await response.json()) as Record<string, unknown>
  const isErrored = Boolean(payload['IsErroredOnProcessing'])
  if (isErrored) return null

  const overlayLines = parseOcrOverlayLines(payload)
  const inferredFromOverlay = inferImportFromOcrOverlay(overlayLines)
  if (inferredFromOverlay && inferredFromOverlay.cards.length > 0) {
    if (inferredFromOverlay.cards.length > 120) return null
    return inferredFromOverlay
  }

  const parsedResults = Array.isArray(payload['ParsedResults'])
    ? (payload['ParsedResults'] as Array<Record<string, unknown>>)
    : []
  const ocrText = parsedResults
    .map((result) => String(result['ParsedText'] ?? ''))
    .join('\n')
    .replace(/\r/g, '\n')
    .trim()

  if (!ocrText) return null
  const parsedImport = parseInputImport(ocrText)
  if (!parsedImport || parsedImport.cards.length === 0) return null

  // OCR noise can explode into dozens of false positives. Keep this bounded.
  if (parsedImport.cards.length > 120) return null

  return {
    cards: parsedImport.cards,
    connections: parsedImport.connections,
    containsRasterImage: true,
  }
}

function normalizeImportCardSpacing(cards: ParsedCardCandidate[]): ParsedCardCandidate[] {
  if (cards.length <= 1) return cards

  const rowThreshold = 70
  const minimumColumnStep = 336
  const sortedByY = [...cards].sort((a, b) => {
    if (Math.abs(a.y - b.y) > 1) return a.y - b.y
    return a.x - b.x
  })

  const rows: ParsedCardCandidate[][] = []
  for (const card of sortedByY) {
    const lastRow = rows[rows.length - 1]
    if (!lastRow) {
      rows.push([card])
      continue
    }
    const rowY = lastRow[0].y
    if (Math.abs(card.y - rowY) <= rowThreshold) {
      lastRow.push(card)
    } else {
      rows.push([card])
    }
  }

  const adjusted = new Map<string, ParsedCardCandidate>()
  for (const row of rows) {
    const sortedRow = [...row].sort((a, b) => a.x - b.x)
    let prevPlacedX: number | null = null
    for (const card of sortedRow) {
      const placedX: number =
        prevPlacedX === null
          ? card.x
          : Math.max(card.x, prevPlacedX + minimumColumnStep)
      adjusted.set(card.sourceNodeId, { ...card, x: placedX })
      prevPlacedX = placedX
    }
  }

  return cards.map((card) => adjusted.get(card.sourceNodeId) ?? card)
}

function isRoleLikeValue(value: string): boolean {
  const normalized = value.toLowerCase().trim()
  return (
    normalized === 'design' ||
    normalized === 'designer' ||
    normalized === 'eng' ||
    normalized === 'engineer' ||
    normalized === 'pm' ||
    normalized === 'ops' ||
    normalized === 'operations' ||
    normalized === 'writing' ||
    normalized === 'research' ||
    normalized === 'data' ||
    normalized.includes('design') ||
    normalized.includes('engineer') ||
    normalized.includes('eng') ||
    normalized.includes('pm') ||
    normalized.includes('ops') ||
    normalized.includes('operations') ||
    normalized.includes('writing') ||
    normalized.includes('research') ||
    normalized.includes('data')
  )
}

function collectTextNodesSorted(node: SceneNode): TextNode[] {
  const textNodes: TextNode[] = []
  if (node.type === 'TEXT') textNodes.push(node)
  if ('findAllWithCriteria' in node) {
    const descendants = node.findAllWithCriteria({ types: ['TEXT'] })
    textNodes.push(...descendants)
  }

  return textNodes
    .filter((textNode) => textNode.characters.trim().length > 0)
    .sort((a, b) => {
      const ay = a.absoluteTransform[1][2]
      const by = b.absoluteTransform[1][2]
      if (Math.abs(ay - by) > 2) return ay - by
      const ax = a.absoluteTransform[0][2]
      const bx = b.absoluteTransform[0][2]
      return ax - bx
    })
}

function nodeHasRasterImage(node: SceneNode): boolean {
  if (node.type === 'STICKY') return false
  if (node.type === 'MEDIA') return true
  if ('fills' in node && Array.isArray(node.fills)) {
    const hasImageFill = node.fills.some((fill) => fill.type === 'IMAGE')
    if (hasImageFill) return true
  }
  if ('findAll' in node) {
    return node.findAll((descendant) => {
      if ('fills' in descendant && Array.isArray(descendant.fills)) {
        return descendant.fills.some((fill) => fill.type === 'IMAGE')
      }
      return descendant.type === 'MEDIA'
    }).length > 0
  }
  return false
}

function deriveProfileFromNode(node: SceneNode): ChipProfile | null {
  const textNodes = collectTextNodesSorted(node)
  const lines = Array.from(
    new Set(
      textNodes
        .map((textNode) => textNode.characters.trim())
        .filter(Boolean),
    ),
  )
  if (lines.length === 0) return null

  let name = lines[0]
  let bestSize = 0
  for (const textNode of textNodes) {
    const size = typeof textNode.fontSize === 'number' ? textNode.fontSize : 0
    if (size >= bestSize && textNode.characters.trim().length > 0) {
      bestSize = size
      name = textNode.characters.trim()
    }
  }

  const roleIndex = lines.findIndex((line) => isRoleLikeValue(line))
  const role = roleIndex >= 0 ? lines[roleIndex] : 'Eng'
  const location =
    roleIndex >= 0 && roleIndex + 1 < lines.length
      ? lines[roleIndex + 1]
      : lines.length >= 3
        ? lines[2]
        : ''

  const projectArea = lines.find((line, index) => {
    if (index === 0) return false
    if (line === role || line === location) return false
    return !isRoleLikeValue(line)
  }) ?? ''

  return {
    name,
    projectArea,
    role: role || 'Eng',
    location,
    status: 'active',
  }
}

function isCardLikeCandidate(node: SceneNode): boolean {
  if (
    node.type === 'CONNECTOR' ||
    node.type === 'WIDGET' ||
    node.type === 'SECTION' ||
    node.type === 'STAMP' ||
    node.type === 'STICKY'
  ) {
    return false
  }
  if (!('width' in node) || !('height' in node)) return false
  if (node.width < 120 || node.width > 700 || node.height < 48 || node.height > 260) return false

  const textCount = collectTextNodesSorted(node).length
  return textCount >= 2
}

function collectCardCandidatesFromSelection(selection: readonly SceneNode[]): SceneNode[] {
  const candidates: SceneNode[] = []

  const addCandidate = (node: SceneNode) => {
    if (!isCardLikeCandidate(node)) return
    if (candidates.some((candidate) => candidate.id === node.id)) return
    candidates.push(node)
  }

  for (const node of selection) {
    addCandidate(node)
    if ('findAll' in node) {
      const descendants = node.findAll((descendant) => isCardLikeCandidate(descendant))
      descendants.forEach(addCandidate)
    }
  }

  // Keep the smallest card-like nodes when nested card-like candidates exist.
  const nestedParentIds = new Set<string>()
  for (const node of candidates) {
    let current: BaseNode | null = node.parent
    while (current) {
      const currentId = current.id
      if (candidates.some((candidate) => candidate.id === currentId)) {
        nestedParentIds.add(currentId)
      }
      current = 'parent' in current ? current.parent : null
    }
  }

  return candidates.filter((node) => !nestedParentIds.has(node.id))
}

async function findOwningCardSourceId(
  endpointNodeId: string,
  sourceCardIds: Set<string>,
): Promise<string | null> {
  if (sourceCardIds.has(endpointNodeId)) return endpointNodeId
  const endpointNode = await figma.getNodeByIdAsync(endpointNodeId)
  if (!endpointNode) return null

  let current: BaseNode | null = endpointNode
  while (current) {
    if (sourceCardIds.has(current.id)) return current.id
    current = 'parent' in current ? current.parent : null
  }
  return null
}

async function parseImportSelection(selection: readonly SceneNode[]): Promise<ImportParseResult> {
  const containsRasterImage = selection.some((node) => nodeHasRasterImage(node))
  const cardNodes = collectCardCandidatesFromSelection(selection)
  const cards: ParsedCardCandidate[] = []

  for (const node of cardNodes) {
    const profile = deriveProfileFromNode(node)
    if (!profile || !profile.name.trim()) continue
    cards.push({
      sourceNodeId: node.id,
      x: node.absoluteTransform[0][2],
      y: node.absoluteTransform[1][2],
      width: node.width,
      height: node.height,
      profile,
    })
  }

  if (cards.length === 0) {
    if (containsRasterImage) {
      return { cards: [], connections: [], containsRasterImage: true }
    }
    const fallbackProfiles = parseSelectionProfiles(
      selection.filter((node) => node.type !== 'WIDGET' && node.type !== 'CONNECTOR'),
    )
    let yOffset = 0
    for (const profile of fallbackProfiles) {
      cards.push({
        sourceNodeId: `fallback-${profile.name}-${yOffset}`,
        x: 0,
        y: yOffset,
        width: PEOPLE_CHIP_WIDTH,
        height: 80,
        profile,
      })
      yOffset += 96
    }
  }

  const connectors: ConnectorNode[] = []
  const connectorIds = new Set<string>()
  for (const node of selection) {
    if (node.type === 'CONNECTOR') {
      connectorIds.add(node.id)
      connectors.push(node)
    }
    if ('findAllWithCriteria' in node) {
      const descendants = node.findAllWithCriteria({ types: ['CONNECTOR'] })
      for (const connector of descendants) {
        if (connectorIds.has(connector.id)) continue
        connectorIds.add(connector.id)
        connectors.push(connector)
      }
    }
  }

  const sourceCardIds = new Set(cards.map((card) => card.sourceNodeId))
  const connections: ParsedConnection[] = []
  const connectionKeys = new Set<string>()

  for (const connector of connectors) {
    const startNodeId = 'endpointNodeId' in connector.connectorStart ? connector.connectorStart.endpointNodeId : null
    const endNodeId = 'endpointNodeId' in connector.connectorEnd ? connector.connectorEnd.endpointNodeId : null
    if (!startNodeId || !endNodeId) continue

    const fromSourceId = await findOwningCardSourceId(startNodeId, sourceCardIds)
    const toSourceId = await findOwningCardSourceId(endNodeId, sourceCardIds)
    if (!fromSourceId || !toSourceId || fromSourceId === toSourceId) continue

    const key = `${fromSourceId}->${toSourceId}`
    if (connectionKeys.has(key)) continue
    connectionKeys.add(key)
    connections.push({ fromSourceId, toSourceId })
  }

  return { cards, connections, containsRasterImage }
}

function InsertionCard(props: {
  label: string
  width: number
  height?: number
  children: FigmaDeclarativeNode
  onClick: () => void
}) {
  return (
    <AutoLayout
      direction={'vertical'}
      width={props.width}
      height={props.height ?? 140}
      fill={'#FFFFFF00'}
      cornerRadius={10}
      horizontalAlignItems={'center'}
      verticalAlignItems={'center'}
      spacing={10}
      onClick={props.onClick}
      hoverStyle={{ fill: '#FFFFFF66' }}
    >
      {props.children}
      <Text fontFamily={'Inter'} fontWeight={500} fontSize={12} fill={'#9B9B9B'} horizontalAlignText={'center'}>
        {props.label}
      </Text>
    </AutoLayout>
  )
}

function PeopleChipPreview() {
  return (
    <AutoLayout
      direction={'horizontal'}
      spacing={16}
      width={CHOOSER_PEOPLE_CARD_WIDTH}
      height={CHOOSER_CARD_HEIGHT}
      padding={16}
      cornerRadius={10}
      fill={'#FFFFFF'}
      stroke={'#CFCFD4'}
      effect={{
        type: 'drop-shadow',
        color: { r: 0, g: 0, b: 0, a: 0.12 },
        offset: { x: 0, y: 2 },
        blur: 4,
        spread: 0,
      }}
      verticalAlignItems={'center'}
    >
      <AutoLayout
        width={48}
        height={48}
        cornerRadius={28}
        overflow={'hidden'}
        verticalAlignItems={'center'}
        horizontalAlignItems={'center'}
        fill={'#F4F4F5'}
        stroke={'#D4D4D8'}
      >
        <Text fontWeight={700} fontSize={17} fill={'#6B7280'}>
          FN
        </Text>
      </AutoLayout>
      <AutoLayout direction={'vertical'} spacing={4} width={PEOPLE_CHIP_DETAILS_WIDTH}>
        <AutoLayout direction={'vertical'} spacing={2}>
          <Text fontWeight={700} fontSize={17} fill={'#111827'} width={PEOPLE_CHIP_DETAILS_WIDTH}>
            First Name
          </Text>
          <Text fontWeight={500} fontSize={13} fill={'#111827'} width={PEOPLE_CHIP_DETAILS_WIDTH}>
            Team Area
          </Text>
        </AutoLayout>
        <AutoLayout direction={'horizontal'} spacing={6} verticalAlignItems={'center'}>
          <AutoLayout
            fill={'#4C8BF5'}
            cornerRadius={4}
            padding={{ left: 4, right: 4, top: 2, bottom: 2 }}
            horizontalAlignItems={'center'}
            verticalAlignItems={'center'}
          >
            <Text fontSize={9} fontWeight={700} fill={'#FFFFFF'}>
              Role
            </Text>
          </AutoLayout>
          <Text fontSize={11} fontWeight={500} fill={'#6B7280'}>
            City
          </Text>
        </AutoLayout>
      </AutoLayout>
    </AutoLayout>
  )
}

function CounterPreview() {
  return (
    <AutoLayout
      width={CHOOSER_SMALL_CARD_WIDTH}
      height={CHOOSER_CARD_HEIGHT}
      horizontalAlignItems={'center'}
      verticalAlignItems={'center'}
      padding={16}
      cornerRadius={10}
      fill={'#FFFFFF'}
      stroke={'#CFCFD4'}
      strokeWidth={1}
      effect={{
        type: 'drop-shadow',
        color: { r: 0, g: 0, b: 0, a: 0.12 },
        offset: { x: 0, y: 2 },
        blur: 4,
      }}
    >
      <Text fontSize={32} fontFamily={'Inter'} fontWeight={500} fill={'#000000'} horizontalAlignText={'center'}>
        8
      </Text>
    </AutoLayout>
  )
}

function BulkCreatePreview() {
  return (
    <AutoLayout
      width={CHOOSER_SMALL_CARD_WIDTH}
      height={CHOOSER_CARD_HEIGHT}
      cornerRadius={10}
      fill={'#FFFFFF'}
      stroke={'#CFCFD4'}
      horizontalAlignItems={'center'}
      verticalAlignItems={'center'}
      effect={{
        type: 'drop-shadow',
        color: { r: 0, g: 0, b: 0, a: 0.12 },
        offset: { x: 0, y: 2 },
        blur: 4,
      }}
    >
      <SVG src={BULK_ICON_DARK} />
    </AutoLayout>
  )
}

function Widget() {
  const widgetNodeId = useWidgetNodeId()
  const [widgetKind, setWidgetKind] = useSyncedState<WidgetKind>('widget-kind', 'chooser')
  const [seed] = useSyncedState<number>('seed', () => Math.floor(Math.random() * 1_000_000))
  const defaultProfile = DEFAULT_CHIP_PROFILES[seed % DEFAULT_CHIP_PROFILES.length]

  const [name, setName] = useSyncedState<string>('name', defaultProfile.name)
  const [projectArea, setProjectArea] = useSyncedState<string>(
    'project-area',
    defaultProfile.projectArea,
  )
  const [role, setRole] = useSyncedState<string>('role', defaultProfile.role)
  const [location, setLocation] = useSyncedState<string>('location', defaultProfile.location)
  const [linkUrl, setLinkUrl] = useSyncedState<string>('link-url', '')
  const [avatarDataUri, setAvatarDataUri] = useSyncedState<string>('avatar-data-uri', '')
  const [status, setStatus] = useSyncedState<Status>('status', 'active')
  const [managerKind, setManagerKind] = useSyncedState<ManagerKind>('manager-kind', 'ic')
  const [count, setCount] = useSyncedState<number | null>('count', null)
  const [counterPreset, setCounterPreset] = useSyncedState<CounterPreset>('counter-preset', 'all')
  const [counterPeopleFilter, setCounterPeopleFilter] = useSyncedState<CounterPeopleFilter>(
    'counter-people-filter',
    'all',
  )
  const [counterBreakdown, setCounterBreakdown] = useSyncedState<string[]>('counter-breakdown', [])
  const [counterRoleDetails, setCounterRoleDetails] = useSyncedState<CounterRoleDetail[]>(
    'counter-role-details',
    [],
  )
  const [bulkInput, setBulkInput] = useSyncedState<string>('bulk-input', '')
  const [bulkSource, setBulkSource] = useSyncedState<BulkSource>('bulk-source', 'text')
  const [bulkSelectionCache, setBulkSelectionCache] = useSyncedState<string>('bulk-selection-cache', '')
  const [bulkSelectionPreviewCount, setBulkSelectionPreviewCount] = useSyncedState<number>(
    'bulk-selection-preview-count',
    0,
  )
  const [bulkSelectionPreviewConnections, setBulkSelectionPreviewConnections] = useSyncedState<number>(
    'bulk-selection-preview-connections',
    0,
  )
  const [bulkSelectionLoading, setBulkSelectionLoading] = useSyncedState<boolean>(
    'bulk-selection-loading',
    false,
  )

  const statusStyle = STATUS_STYLES[status]
  const rolePillColor = roleColor(role)
  const nameFieldWidth = textFieldWidth(name, 'Name', 120, 156, 8.2)
  const nameFieldFontSize = fittedFontSizeForWidth(name, 'Name', nameFieldWidth, 17, 13)
  const projectAreaFieldWidth = textFieldWidth(projectArea, 'Project Area', 112, 156, 6.3)
  const locationFieldWidth = textFieldWidth(location, 'Location', 78, 116, 5.8)
  const counterCardStyle =
    counterPreset === 'all' || counterPreset === 'locations'
      ? STATUS_STYLES.active
      : STATUS_STYLES[counterPreset]
  const counterRefreshStyle =
    counterPreset === 'new-headcount'
      ? { fill: '#EDE1B9', hover: '#E4D7AD', icon: '#B7A96F' }
      : counterPreset === 'open-headcount'
        ? { fill: '#D9D9DD', hover: '#CFD0D4', icon: '#A7A8AE' }
        : { fill: '#F1F1F1', hover: '#E3E3E3', icon: '#B0B0B4' }

  const uploadAvatarImage = async () => {
    figma.showUI(AVATAR_UPLOAD_UI_HTML, {
      width: 360,
      height: 156,
      title: 'Upload profile image',
    })
    figma.ui.postMessage({ type: 'avatar-upload-init', hasImage: avatarDataUri.trim().length > 0 })

    await new Promise<void>((resolve) => {
      let done = false
      const finish = () => {
        if (done) return
        done = true
        figma.ui.onmessage = undefined
        figma.ui.close()
        resolve()
      }

      const timeoutId = setTimeout(() => {
        figma.notify('Image upload timed out. Try again.')
        finish()
      }, AVATAR_UPLOAD_TIMEOUT_MS)

      figma.ui.onmessage = (message: unknown) => {
        const parsed = asAvatarUploadMessage(message)
        if (!parsed) return
        clearTimeout(timeoutId)

        if (parsed.type === 'avatar-uploaded') {
          if (!parsed.dataUri) {
            figma.notify('Could not upload image.')
            finish()
            return
          }
          setAvatarDataUri(parsed.dataUri)
          finish()
          return
        }

        if (parsed.type === 'avatar-upload-error') {
          figma.notify(parsed.message)
          finish()
          return
        }

        if (parsed.type === 'avatar-remove-image') {
          setAvatarDataUri('')
          finish()
          return
        }

        finish()
      }
    })
  }

  const editChipLink = async () => {
    figma.showUI(LINK_EDITOR_UI_HTML, {
      width: 360,
      height: 150,
      title: 'People chip link',
    })
    figma.ui.postMessage({ type: 'link-editor-init', url: linkUrl })

    await new Promise<void>((resolve) => {
      let done = false
      const finish = () => {
        if (done) return
        done = true
        figma.ui.onmessage = undefined
        figma.ui.close()
        resolve()
      }

      figma.ui.onmessage = (message: unknown) => {
        const parsed = asLinkEditMessage(message)
        if (!parsed) return
        if (parsed.type === 'link-cancel') {
          finish()
          return
        }
        const normalized = normalizeExternalUrl(parsed.url)
        setLinkUrl(normalized)
        finish()
      }
    })
  }

  const refreshCount = async (presetOverride?: CounterPreset, peopleFilterOverride?: CounterPeopleFilter) => {
    const effectivePreset = presetOverride ?? counterPreset
    const effectivePeopleFilter = peopleFilterOverride ?? counterPeopleFilter
    const widgetNode = await figma.getNodeByIdAsync(widgetNodeId)
    if (!widgetNode || widgetNode.type !== 'WIDGET') return

    const containingSection = findContainingSectionByGeometry(widgetNode)
    const allChips = containingSection
      ? getPeopleChipWidgetsInSection(containingSection, widgetNode.widgetId)
      : (figma.currentPage.findAll((node) => isPeopleChipWidgetNode(node, widgetNode.widgetId)) as WidgetNode[])
    const filteredChips = allChips.filter((chip) => matchesPeopleFilter(chip, effectivePeopleFilter))
    const matchingChips = filteredChips.filter((chip) => matchesPreset(chip, effectivePreset))
    const nextBreakdown = buildRoleBreakdown(matchingChips)
    const nextRoleDetails =
      effectivePreset === 'all'
        ? buildCounterRoleDetails(matchingChips)
        : effectivePreset === 'locations'
          ? buildCounterLocationDetails(matchingChips)
          : buildCounterRoleDetails(matchingChips)
    const nextCount = effectivePreset === 'locations' ? nextRoleDetails.length : matchingChips.length
    if (nextCount !== count) {
      setCount(nextCount)
    }
    if (JSON.stringify(nextBreakdown) !== JSON.stringify(counterBreakdown)) {
      setCounterBreakdown(nextBreakdown)
    }
    if (JSON.stringify(nextRoleDetails) !== JSON.stringify(counterRoleDetails)) {
      setCounterRoleDetails(nextRoleDetails)
    }
  }

  const spawnNewWidgetInstance = async () => {
    const widgetNode = await figma.getNodeByIdAsync(widgetNodeId)
    if (!widgetNode || widgetNode.type !== 'WIDGET') return

    const newSeed = Math.floor(Math.random() * 1_000_000)
    const nextProfile = DEFAULT_CHIP_PROFILES[newSeed % DEFAULT_CHIP_PROFILES.length]
    const sourceTopLeft = getTopLeftPoint(widgetNode)
    const targetAbsoluteX = sourceTopLeft.x + widgetNode.width + 24
    const targetAbsoluteY = sourceTopLeft.y

    const clone = widgetNode.cloneWidget({
      'widget-kind': 'chooser',
      seed: newSeed,
      name: nextProfile.name,
      'project-area': nextProfile.projectArea,
      role: nextProfile.role,
      location: nextProfile.location,
      'link-url': '',
      'avatar-data-uri': '',
      status: 'active',
      'manager-kind': 'ic',
      count: null,
      'counter-breakdown': [],
      'counter-role-details': [],
      'counter-preset': 'all',
      'counter-people-filter': 'all',
    })

    const sourceParent = widgetNode.parent
    if (sourceParent && 'appendChild' in sourceParent) {
      sourceParent.appendChild(clone)
      const parentTopLeft = getParentTopLeft(sourceParent)
      clone.x = targetAbsoluteX - parentTopLeft.x
      clone.y = targetAbsoluteY - parentTopLeft.y
    } else {
      clone.x = targetAbsoluteX
      clone.y = targetAbsoluteY
    }

    figma.currentPage.selection = [clone]
    figma.viewport.scrollAndZoomIntoView([clone])
  }

  const createPeopleChipWidgets = async (profiles: ChipProfile[]) => {
    const widgetNode = await figma.getNodeByIdAsync(widgetNodeId)
    if (!widgetNode || widgetNode.type !== 'WIDGET') return
    if (profiles.length === 0) return

    const sourceTopLeft = getTopLeftPoint(widgetNode)
    const startAbsoluteX = sourceTopLeft.x
    const startAbsoluteY = sourceTopLeft.y + widgetNode.height + 20
    const columnWidth = 320
    const rowHeight = 96
    const created: WidgetNode[] = []

    const sourceParent = widgetNode.parent
    const parentTopLeft =
      sourceParent && 'appendChild' in sourceParent ? getParentTopLeft(sourceParent) : { x: 0, y: 0 }

    for (let i = 0; i < profiles.length; i += 1) {
      const profile = profiles[i]
      const clone = widgetNode.cloneWidget({
        'widget-kind': 'people-chip',
        name: profile.name,
        'project-area': profile.projectArea,
        role: profile.role || 'Eng',
        location: profile.location,
        'link-url': '',
        'avatar-data-uri': '',
        status: profile.status,
        'manager-kind': 'ic',
      })
      if (sourceParent && 'appendChild' in sourceParent) {
        sourceParent.appendChild(clone)
      }

      const row = Math.floor(i / 2)
      const col = i % 2
      const x = startAbsoluteX + col * columnWidth
      const y = startAbsoluteY + row * rowHeight
      clone.x = x - parentTopLeft.x
      clone.y = y - parentTopLeft.y
      created.push(clone)
    }

    figma.currentPage.selection = created
    figma.viewport.scrollAndZoomIntoView(created)
  }

  const createPeopleChipWidgetsFromImport = async (parsed: ImportParseResult) => {
    const widgetNode = await figma.getNodeByIdAsync(widgetNodeId)
    if (!widgetNode || widgetNode.type !== 'WIDGET') return
    if (parsed.cards.length === 0) return
    const spacedCards = normalizeImportCardSpacing(parsed.cards)

    const sourceParent = widgetNode.parent
    const parentTopLeft =
      sourceParent && 'appendChild' in sourceParent ? getParentTopLeft(sourceParent) : { x: 0, y: 0 }

    const minX = Math.min(...spacedCards.map((card) => card.x))
    const minY = Math.min(...spacedCards.map((card) => card.y))
    const sourceTopLeft = getTopLeftPoint(widgetNode)
    const baseAbsoluteX = sourceTopLeft.x
    const baseAbsoluteY = sourceTopLeft.y + widgetNode.height + 20

    const createdBySourceId = new Map<string, WidgetNode>()
    const createdNodes: WidgetNode[] = []

    for (const card of spacedCards) {
      const clone = widgetNode.cloneWidget({
        'widget-kind': 'people-chip',
        name: card.profile.name,
        'project-area': card.profile.projectArea,
        role: card.profile.role || 'Eng',
        location: card.profile.location,
        'link-url': '',
        'avatar-data-uri': '',
        status: card.profile.status,
        'manager-kind': 'ic',
      })

      if (sourceParent && 'appendChild' in sourceParent) {
        sourceParent.appendChild(clone)
      }

      const absoluteX = baseAbsoluteX + (card.x - minX)
      const absoluteY = baseAbsoluteY + (card.y - minY)
      clone.x = absoluteX - parentTopLeft.x
      clone.y = absoluteY - parentTopLeft.y

      createdBySourceId.set(card.sourceNodeId, clone)
      createdNodes.push(clone)
    }

    const createdConnectors: ConnectorNode[] = []
    for (const connection of parsed.connections) {
      const fromNode = createdBySourceId.get(connection.fromSourceId)
      const toNode = createdBySourceId.get(connection.toSourceId)
      if (!fromNode || !toNode) continue

      try {
        const connector = figma.createConnector()
        connector.connectorLineType = 'ELBOWED'
        connector.strokes = [{ type: 'SOLID', color: { r: 0.67, g: 0.67, b: 0.7 } }]
        const fromY = fromNode.absoluteTransform[1][2]
        const toY = toNode.absoluteTransform[1][2]
        const downward = toY >= fromY

        connector.connectorStart = {
          endpointNodeId: fromNode.id,
          magnet: downward ? 'BOTTOM' : 'TOP',
        }
        connector.connectorEnd = {
          endpointNodeId: toNode.id,
          magnet: downward ? 'TOP' : 'BOTTOM',
        }
        connector.connectorEndStrokeCap = 'ARROW_LINES'
        if (sourceParent && 'appendChild' in sourceParent) {
          sourceParent.appendChild(connector)
        }
        createdConnectors.push(connector)
      } catch {
        // Connector creation is not available in all editor contexts.
      }
    }

    figma.currentPage.selection = [...createdNodes, ...createdConnectors]
    figma.viewport.scrollAndZoomIntoView(createdNodes)
  }

  const runBulkCreate = async () => {
    if (bulkSelectionLoading) return
    if (bulkSource === 'text') {
      const inputImport = parseInputImport(bulkInput)
      if (inputImport && inputImport.cards.length > 0) {
        await createPeopleChipWidgetsFromImport(inputImport)
        figma.notify(
          `Created ${inputImport.cards.length} cards${inputImport.connections.length > 0 ? ` and ${inputImport.connections.length} connectors` : ''}.`,
        )
        return
      }
      figma.notify('Paste valid text/JSON first.')
      return
    }

    if (bulkSelectionCache.trim().length > 0) {
      try {
        const parsedSelection = JSON.parse(bulkSelectionCache) as ImportParseResult
        if (parsedSelection.cards.length > 0) {
          await createPeopleChipWidgetsFromImport(parsedSelection)
          figma.notify(
            `Created ${parsedSelection.cards.length} cards${parsedSelection.connections.length > 0 ? ` and ${parsedSelection.connections.length} connectors` : ''}.`,
          )
          return
        }
      } catch {
        // Selection cache parse failed; fall through to guidance toast.
      }
    }

    figma.notify('Select and preview canvas frames first.')
  }

  const runSelectionPreview = async () => {
    if (bulkSelectionLoading) return
    setBulkSelectionLoading(true)
    try {
      const parsedSelection = await parseImportSelection(figma.currentPage.selection)
      const looksLikeFallbackSelection =
        parsedSelection.cards.length > 0 &&
        parsedSelection.connections.length === 0 &&
        parsedSelection.cards.every((card) => {
          const name = card.profile.name.trim().toLowerCase()
          return (
            card.sourceNodeId.startsWith('fallback-') ||
            /^image\s*\d*$/.test(name) ||
            /^frame\s*\d*$/.test(name) ||
            /^rectangle\s*\d*$/.test(name)
          )
        })

      if (parsedSelection.containsRasterImage && (parsedSelection.cards.length === 0 || looksLikeFallbackSelection)) {
        const ocrParsed = await parseRasterSelectionWithOcr(figma.currentPage.selection)
        if (ocrParsed && ocrParsed.cards.length > 0) {
          setBulkSelectionCache(JSON.stringify(ocrParsed))
          setBulkSelectionPreviewCount(ocrParsed.cards.length)
          setBulkSelectionPreviewConnections(ocrParsed.connections.length)
          setBulkSource('selection')
          figma.notify(
            `OCR detected ${ocrParsed.cards.length} ${pluralize(ocrParsed.cards.length, 'person', 'people')}${ocrParsed.connections.length > 0 ? ` and ${ocrParsed.connections.length} ${pluralize(ocrParsed.connections.length, 'connection', 'connections')}` : ''}.`,
          )
          return
        }
      }

      if (parsedSelection.cards.length > 0) {
        setBulkSelectionCache(JSON.stringify(parsedSelection))
        setBulkSelectionPreviewCount(parsedSelection.cards.length)
        setBulkSelectionPreviewConnections(parsedSelection.connections.length)
        setBulkSource('selection')
        return
      }
      setBulkSelectionCache('')
      setBulkSelectionPreviewCount(0)
      setBulkSelectionPreviewConnections(0)
      if (parsedSelection.containsRasterImage) {
        const ocrParsed = await parseRasterSelectionWithOcr(figma.currentPage.selection)
        if (ocrParsed && ocrParsed.cards.length > 0) {
          setBulkSelectionCache(JSON.stringify(ocrParsed))
          setBulkSelectionPreviewCount(ocrParsed.cards.length)
          setBulkSelectionPreviewConnections(ocrParsed.connections.length)
          setBulkSource('selection')
          figma.notify(
            `OCR detected ${ocrParsed.cards.length} ${pluralize(ocrParsed.cards.length, 'person', 'people')}${ocrParsed.connections.length > 0 ? ` and ${ocrParsed.connections.length} ${pluralize(ocrParsed.connections.length, 'connection', 'connections')}` : ''}.`,
          )
          return
        }
        figma.notify('Could not interpret this image yet. Try selecting a tighter crop, or use text/JSON input.')
        return
      }
      figma.notify('Select card-like frames/connectors on the canvas, then try again.')
    } catch (error) {
      const message = error instanceof Error ? error.message : 'OCR failed.'
      figma.notify(message)
    } finally {
      setBulkSelectionLoading(false)
    }
  }

  const parsedInputImport = parseInputImport(bulkInput)
  const canCreateFromInput = Boolean(parsedInputImport && parsedInputImport.cards.length > 0)
  const canCreateFromSelection = bulkSelectionCache.trim().length > 0
  const canBulkCreate = bulkSource === 'text' ? !bulkSelectionLoading : canCreateFromSelection && !bulkSelectionLoading
  const bulkButtonHeight = 44
  const bulkStatusMessage =
    bulkSelectionLoading
      ? 'Scanning selection with OCR...'
      : bulkSource === 'text'
      ? bulkInput.trim().length > 0
        ? canCreateFromInput
          ? detectedSummary(parsedInputImport?.cards.length ?? 0, parsedInputImport?.connections.length ?? 0)
          : 'Click Create cards to try import'
        : ''
      : canCreateFromSelection
        ? detectedSummary(bulkSelectionPreviewCount, bulkSelectionPreviewConnections)
        : ''

  const propertyMenuItems: WidgetPropertyMenuItem[] =
    widgetKind === 'people-chip'
      ? [
          {
            itemType: 'dropdown',
            propertyName: 'status',
            tooltip: '',
            selectedOption: status,
            options: STATUS_OPTIONS,
          },
          {
            itemType: 'dropdown',
            propertyName: 'role',
            tooltip: '',
            selectedOption: role,
            options: ROLE_OPTIONS,
          },
          {
            itemType: 'action',
            propertyName: 'manager-kind-toggle',
            tooltip: '',
            icon: managerToggleIcon(managerKind),
          },
          { itemType: 'separator' },
          {
            itemType: 'action',
            propertyName: 'edit-link',
            tooltip: '',
            icon: LINK_ICON,
          },
          ...(linkUrl.trim().length > 0
            ? [
                {
                  itemType: 'action',
                  propertyName: 'open-link',
                  tooltip: '',
                  icon: OPEN_LINK_ICON,
                } satisfies WidgetPropertyMenuItem,
              ]
            : []),
        ]
      : widgetKind === 'counter'
        ? [
            {
              itemType: 'dropdown',
              propertyName: 'counter-preset',
              tooltip: '',
              selectedOption: counterPreset,
              options: COUNTER_PRESET_OPTIONS,
            },
            {
              itemType: 'dropdown',
              propertyName: 'counter-people-filter',
              tooltip: '',
              selectedOption: counterPeopleFilter,
              options: COUNTER_PEOPLE_FILTER_OPTIONS,
            },
          ]
        : []

  usePropertyMenu(
    propertyMenuItems,
    async (event) => {
      if (event.propertyName === 'status') {
        const nextStatus = event.propertyValue as Status | undefined
        if (nextStatus && STATUS_STYLES[nextStatus]) {
          setStatus(nextStatus)
        }
        return
      }
      if (event.propertyName === 'role') {
        const nextRole = event.propertyValue as string | undefined
        if (nextRole) {
          setRole(nextRole)
        }
        return
      }
      if (event.propertyName === 'counter-preset') {
        const nextPreset = event.propertyValue as CounterPreset | undefined
        if (nextPreset && COUNTER_PRESET_LABEL[nextPreset]) {
          setCounterPreset(nextPreset)
          await refreshCount(nextPreset)
        }
        return
      }
      if (event.propertyName === 'counter-people-filter') {
        const nextPeopleFilter = event.propertyValue as CounterPeopleFilter | undefined
        if (nextPeopleFilter) {
          setCounterPeopleFilter(nextPeopleFilter)
          await refreshCount(undefined, nextPeopleFilter)
        }
        return
      }
      if (event.propertyName === 'manager-kind-toggle') {
        setManagerKind(managerKind === 'manager' ? 'ic' : 'manager')
        return
      }
      if (event.propertyName === 'edit-link') {
        await editChipLink()
        return
      }
      if (event.propertyName === 'open-link') {
        const url = normalizeExternalUrl(linkUrl)
        if (!url) return
        figma.openExternal(url)
        return
      }
    },
  )

  if (widgetKind === 'chooser') {
    return (
      <AutoLayout
        key={'mode-chooser'}
        direction={'vertical'}
        spacing={12}
        horizontalAlignItems={'center'}
        padding={{ top: 24, right: 8, bottom: 8, left: 8 }}
        cornerRadius={10}
        stroke={'#C4C4C4'}
        fill={'#F5F5F5'}
      >
        <Text fontFamily={'Inter'} fontWeight={500} fontSize={15} fill={'#9B9B9B'}>
          Choose a widget
        </Text>
        <AutoLayout
          direction={'vertical'}
          spacing={0}
          height={260}
          horizontalAlignItems={'center'}
        >
          <InsertionCard
            label={'Add people chip'}
            width={280}
            height={120}
            onClick={() => {
              setWidgetKind('people-chip')
            }}
          >
            <PeopleChipPreview />
          </InsertionCard>
          <AutoLayout direction={'horizontal'} spacing={16} width={328} horizontalAlignItems={'center'} verticalAlignItems={'start'}>
            <InsertionCard
              label={'Add counter'}
              width={132}
              height={120}
              onClick={() => {
                setWidgetKind('counter')
                void refreshCount()
              }}
            >
              <CounterPreview />
            </InsertionCard>
            <InsertionCard
              label={'Bulk create'}
              width={132}
              height={120}
              onClick={() => {
                setWidgetKind('bulk-create')
              }}
            >
              <BulkCreatePreview />
            </InsertionCard>
          </AutoLayout>
        </AutoLayout>
      </AutoLayout>
    )
  }

  if (widgetKind === 'bulk-create') {
    return (
      <AutoLayout
        key={'mode-bulk-create'}
        direction={'vertical'}
        spacing={18}
        padding={{ top: 22, right: 28, bottom: 20, left: 28 }}
        cornerRadius={10}
        stroke={'#C4C4C4'}
        fill={'#F5F5F5'}
        width={588}
      >
        <AutoLayout
          direction={'horizontal'}
          spacing={12}
          verticalAlignItems={'center'}
          onClick={() => {
            setWidgetKind('chooser')
          }}
        >
          <Text fontFamily={'Inter'} fontWeight={500} fontSize={20} fill={'#9B9B9B'}>
            ←
          </Text>
          <Text fontFamily={'Inter'} fontWeight={500} fontSize={17} fill={'#9B9B9B'}>
            Bulk create
          </Text>
        </AutoLayout>

        <AutoLayout direction={'horizontal'} spacing={18} verticalAlignItems={'center'}>
          <Text
            fontFamily={'Inter'}
            fontWeight={bulkSource === 'text' ? 600 : 500}
            fontSize={15}
            fill={bulkSource === 'text' ? '#8D46FF' : '#9B9B9B'}
            onClick={() => {
              setBulkSource('text')
            }}
          >
            Text/JSON
          </Text>
          <Text
            fontFamily={'Inter'}
            fontWeight={bulkSource === 'selection' ? 600 : 500}
            fontSize={15}
            fill={bulkSource === 'selection' ? '#8D46FF' : '#9B9B9B'}
            onClick={() => {
              setBulkSource('selection')
            }}
          >
            Canvas selection
          </Text>
        </AutoLayout>

        {bulkSource === 'text' ? (
          <AutoLayout direction={'horizontal'} spacing={12} verticalAlignItems={'center'}>
            <Input
              name={'OrgJam/BulkInput'}
              value={bulkInput}
              onTextEditEnd={(e) => {
                setBulkInput(e.characters || '')
              }}
              inputFrameProps={{
                name: 'OrgJam/BulkInputField',
                fill: '#FFFFFF',
                stroke: '#CFCFD4',
                cornerRadius: 10,
                padding: { left: 16, right: 16, top: 13, bottom: 13 },
              }}
              width={370}
              fontFamily={'Inter'}
              fontSize={14}
              fontWeight={500}
              fill={bulkInput.trim().length > 0 ? '#6F6F75' : '#9B9B9B'}
              inputBehavior={'multiline'}
              placeholder={'Paste text or JSON'}
            />
            <AutoLayout
              width={150}
              height={bulkButtonHeight}
              cornerRadius={10}
              fill={canBulkCreate ? '#8D46FF' : '#CDB3FF'}
              horizontalAlignItems={'center'}
              verticalAlignItems={'center'}
              onClick={
                canBulkCreate
                  ? () => {
                      waitForTask(runBulkCreate())
                    }
                  : undefined
              }
              hoverStyle={canBulkCreate ? { fill: '#7C37F0' } : undefined}
            >
              <Text fontFamily={'Inter'} fontWeight={600} fontSize={15} fill={'#FFFFFF'}>
                Create cards
              </Text>
            </AutoLayout>
          </AutoLayout>
        ) : (
          <AutoLayout direction={'horizontal'} spacing={12} verticalAlignItems={'center'}>
            <AutoLayout
              width={370}
              height={44}
              cornerRadius={10}
              fill={'#FFFFFF'}
              stroke={'#CFCFD4'}
              horizontalAlignItems={'center'}
              verticalAlignItems={'center'}
              onClick={() => {
                waitForTask(runSelectionPreview())
              }}
              hoverStyle={bulkSelectionLoading ? undefined : { fill: '#F7F7F8' }}
            >
              {bulkSelectionLoading ? (
                <AutoLayout direction={'horizontal'} spacing={8} verticalAlignItems={'center'}>
                  <SVG src={refreshIconSvg('#8D46FF')} />
                  <Text fontFamily={'Inter'} fontWeight={500} fontSize={15} fill={'#8D46FF'}>
                    Scanning selection...
                  </Text>
                </AutoLayout>
              ) : (
                <Text fontFamily={'Inter'} fontWeight={500} fontSize={15} fill={'#6B6B72'}>
                  {canCreateFromSelection ? 'Frames selected' : 'Make a selection then click here'}
                </Text>
              )}
            </AutoLayout>
            <AutoLayout
              width={150}
              height={bulkButtonHeight}
              cornerRadius={10}
              fill={canBulkCreate ? '#8D46FF' : '#CDB3FF'}
              horizontalAlignItems={'center'}
              verticalAlignItems={'center'}
              onClick={
                canBulkCreate
                  ? () => {
                      waitForTask(runBulkCreate())
                    }
                  : undefined
              }
              hoverStyle={canBulkCreate ? { fill: '#7C37F0' } : undefined}
            >
              <Text fontFamily={'Inter'} fontWeight={600} fontSize={15} fill={'#FFFFFF'}>
                Create cards
              </Text>
            </AutoLayout>
          </AutoLayout>
        )}
        {bulkStatusMessage ? (
          <Text fontFamily={'Inter'} fontWeight={500} fontSize={12} fill={'#9B9B9B'}>
            {bulkStatusMessage}
          </Text>
        ) : null}
      </AutoLayout>
    )
  }

  if (widgetKind === 'counter') {
    return (
      <AutoLayout
        key={'mode-counter'}
        name={'headCount/Frame'}
        width={PEOPLE_CHIP_WIDTH}
        direction={'horizontal'}
        spacing={count === null ? 24 : 0}
        horizontalAlignItems={'center'}
        verticalAlignItems={count === null ? 'center' : 'start'}
        padding={{ top: 20, right: 27, bottom: 20, left: 27 }}
        cornerRadius={10}
        fill={counterCardStyle.cardFill}
        stroke={counterCardStyle.cardStroke}
        strokeWidth={1}
        effect={{
          type: 'drop-shadow',
          color: { r: 0, g: 0, b: 0, a: 0.12 },
          offset: { x: 0, y: 2 },
          blur: 4,
        }}
      >
        {count === null ? (
          <AutoLayout
            direction={'horizontal'}
            spacing={12}
            verticalAlignItems={'center'}
          >
            <Text fontSize={17} fontFamily={'Inter'} fontWeight={500} fill={'#7A7A7A'} width={150}>
              Click refresh
              {'\n'}
              to count
            </Text>
            <AutoLayout
              name={'headCount/Refresh'}
              width={44}
              height={40}
              cornerRadius={8}
              fill={counterRefreshStyle.fill}
              verticalAlignItems={'center'}
              horizontalAlignItems={'center'}
              onClick={() => {
                void refreshCount()
              }}
              hoverStyle={{ fill: counterRefreshStyle.hover }}
            >
              <SVG src={refreshIconSvg(counterRefreshStyle.icon)} />
            </AutoLayout>
          </AutoLayout>
        ) : (
          <>
            <AutoLayout direction={'vertical'} spacing={15} width={COUNTER_DETAILS_WIDTH}>
              <AutoLayout direction={'vertical'} spacing={2}>
                <Text fontSize={8} fontFamily={'Inter'} fontWeight={500} fill={'#8A8A8F'} letterSpacing={1}>
                  {COUNTER_PRESET_LABEL[counterPreset].toUpperCase()}
                </Text>
                <Text fontSize={32} fontFamily={'Inter'} fontWeight={500} fill={'#000000'} horizontalAlignText={'center'}>
                  {count}
                </Text>
              </AutoLayout>

              {counterRoleDetails.length > 0 ? (
                <AutoLayout direction={'vertical'} spacing={14} width={COUNTER_DETAILS_WIDTH}>
                  {counterRoleDetails.map((detail, index) => (
                    <AutoLayout key={`${detail.title}-${index}`} direction={'vertical'} spacing={1} width={COUNTER_DETAILS_WIDTH}>
                      <Text fontSize={12} fontFamily={'Inter'} fontWeight={700} fill={detail.color} width={COUNTER_DETAILS_WIDTH}>
                        {detail.title}
                      </Text>
                      {counterPreset === 'all' ? (
                        <Text fontSize={10} fontFamily={'Inter'} fontWeight={500} fill={'#8A8A8F'} width={COUNTER_DETAILS_WIDTH}>
                          {detail.statusSummary}
                        </Text>
                      ) : null}
                    </AutoLayout>
                  ))}
                </AutoLayout>
              ) : counterBreakdown.length > 0 ? (
                <AutoLayout direction={'vertical'} spacing={1}>
                  {counterBreakdown.map((line, index) => (
                    <Text key={`${line}-${index}`} fontSize={11} fontFamily={'Inter'} fontWeight={500} fill={'#7A7A7A'}>
                      {line}
                    </Text>
                  ))}
                </AutoLayout>
              ) : null}
            </AutoLayout>
            <AutoLayout padding={{ top: 12 }}>
              <AutoLayout
                name={'headCount/Refresh'}
                width={44}
                height={40}
                cornerRadius={8}
                fill={counterRefreshStyle.fill}
                verticalAlignItems={'center'}
                horizontalAlignItems={'center'}
                onClick={() => {
                  void refreshCount()
                }}
                hoverStyle={{ fill: counterRefreshStyle.hover }}
              >
                <SVG src={refreshIconSvg(counterRefreshStyle.icon)} />
              </AutoLayout>
            </AutoLayout>
          </>
        )}
      </AutoLayout>
    )
  }

  return (
    <AutoLayout
      key={'mode-people-chip'}
      name={'OrgJam/PeopleChip'}
      direction={'horizontal'}
      spacing={16}
      width={PEOPLE_CHIP_WIDTH}
      padding={16}
      cornerRadius={10}
      fill={statusStyle.cardFill}
      stroke={statusStyle.cardStroke}
      effect={{
        type: 'drop-shadow',
        color: { r: 0, g: 0, b: 0, a: 0.12 },
        offset: { x: 0, y: 2 },
        blur: 4,
        spread: 0,
      }}
      verticalAlignItems={'center'}
    >
      <AutoLayout
        name={'OrgJam/Avatar'}
        width={48}
        height={48}
        cornerRadius={28}
        overflow={'hidden'}
        verticalAlignItems={'center'}
        horizontalAlignItems={'center'}
        fill={statusStyle.avatarFill}
        stroke={statusStyle.avatarStroke}
        onClick={() => {
          waitForTask(uploadAvatarImage())
        }}
      >
        {avatarDataUri ? (
          <Image
            name={'OrgJam/AvatarImage'}
            src={avatarDataUri}
            width={48}
            height={48}
            cornerRadius={24}
          />
        ) : (
          <Text
            name={'OrgJam/AvatarInitials'}
            fontWeight={status === 'active' ? 700 : 500}
            fontSize={18}
            fill={statusStyle.avatarText}
          >
            {initialsFromName(name)}
          </Text>
        )}
      </AutoLayout>

      <AutoLayout name={'OrgJam/Details'} direction={'vertical'} spacing={4} width={PEOPLE_CHIP_DETAILS_WIDTH}>
        <AutoLayout name={'OrgJam/PrimaryText'} direction={'vertical'} spacing={2}>
          <Input
            name={'OrgJam/Name'}
            value={name}
            onTextEditEnd={(e) => setName(e.characters || '')}
            inputFrameProps={{ name: 'OrgJam/NameField' }}
            fontWeight={status === 'active' ? 700 : 500}
            fontSize={nameFieldFontSize}
            fill={'#111827'}
            inputBehavior={'truncate'}
            placeholder={'Name'}
            width={nameFieldWidth}
          />

          <Input
            name={'OrgJam/ProjectArea'}
            value={projectArea}
            onTextEditEnd={(e) => setProjectArea(e.characters || '')}
            inputFrameProps={{ name: 'OrgJam/ProjectAreaField' }}
            fontWeight={500}
            fontSize={13}
            fill={'#111827'}
            inputBehavior={'truncate'}
            placeholder={'Project Area'}
            width={projectAreaFieldWidth}
          />
        </AutoLayout>

        <AutoLayout
          name={'OrgJam/MetaRow'}
          direction={'horizontal'}
          spacing={6}
          verticalAlignItems={'center'}
        >
          <AutoLayout
            name={'OrgJam/RolePill'}
            fill={rolePillColor}
            cornerRadius={4}
            padding={{ left: 4, right: 4, top: 2, bottom: 2 }}
            horizontalAlignItems={'center'}
            verticalAlignItems={'center'}
          >
            <Input
              name={'OrgJam/Role'}
              value={role}
              onTextEditEnd={(e) => setRole(e.characters || '')}
              inputFrameProps={{ name: 'OrgJam/RoleField' }}
              fontSize={9}
              fontWeight={700}
              fill={'#FFFFFF'}
              inputBehavior={'truncate'}
              horizontalAlignText={'center'}
              placeholder={'Role'}
              width={roleFieldWidth(role)}
            />
          </AutoLayout>

          <Input
            name={'OrgJam/Location'}
            value={location}
            onTextEditEnd={(e) => setLocation(e.characters || '')}
            inputFrameProps={{ name: 'OrgJam/LocationField' }}
            fontSize={11}
            fontWeight={500}
            fill={'#6B7280'}
            inputBehavior={'truncate'}
            placeholder={'Location'}
            width={locationFieldWidth}
          />
        </AutoLayout>

      </AutoLayout>
    </AutoLayout>
  )
}

widget.register(Widget)
