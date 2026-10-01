import { FormEvent, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

type Applicant = {
  id: number
  name: string
  department: string
  grade: string
  email: string
  phone: string
  startDate: string
  activity: string
  days: string[]
  startTime: string
  endTime: string
  roles: string[]
  interest: string
  memo: string
  createdAt: string
}

const STORAGE_KEY = 'campusMateApplicants'
const days = ['월', '화', '수', '목', '금', '토', '일']
const roles = ['발표', '자료조사', 'PPT', '문서작성', '코딩', '아이디어']

function loadApplicants(): Applicant[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEY)
    return saved ? JSON.parse(saved) : []
  } catch {
    return []
  }
}

function App() {
  const [applicants, setApplicants] = useState<Applicant[]>(loadApplicants)
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null)
  const [search, setSearch] = useState('')
  const [activityFilter, setActivityFilter] = useState('all')
  const [dayFilter, setDayFilter] = useState('all')
  const formRef = useRef<HTMLFormElement>(null)
  const nameRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(applicants))
  }, [applicants])

  const stats = useMemo(() => ({
    total: applicants.length,
    team: applicants.filter((p) => p.activity === '팀플' || p.activity === '둘 다 가능').length,
    study: applicants.filter((p) => p.activity === '스터디' || p.activity === '둘 다 가능').length,
  }), [applicants])

  const filteredApplicants = useMemo(() => {
    const keyword = search.trim().toLowerCase()
    return applicants.filter((person) => {
      const text = [
        person.name,
        person.department,
        person.grade,
        person.activity,
        person.interest,
        person.email,
        person.phone,
        person.memo,
        ...person.days,
        ...person.roles,
      ].join(' ').toLowerCase()

      const matchesKeyword = text.includes(keyword)
      const matchesActivity = activityFilter === 'all' || person.activity === activityFilter
      const matchesDay = dayFilter === 'all' || person.days.includes(dayFilter)
      return matchesKeyword && matchesActivity && matchesDay
    })
  }, [applicants, search, activityFilter, dayFilter])

  const showMessage = (text: string, type: 'success' | 'error' = 'success') => {
    setMessage({ text, type })
  }

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setMessage(null)

    const data = new FormData(event.currentTarget)
    const name = String(data.get('name') ?? '').trim()
    const department = String(data.get('department') ?? '').trim()
    const grade = String(data.get('grade') ?? '')
    const email = String(data.get('email') ?? '').trim()
    const phone = String(data.get('phone') ?? '').trim()
    const startDate = String(data.get('startDate') ?? '')
    const activity = String(data.get('activity') ?? '')
    const selectedDays = data.getAll('days').map(String)
    const startTime = String(data.get('startTime') ?? '')
    const endTime = String(data.get('endTime') ?? '')
    const selectedRoles = data.getAll('roles').map(String)
    const interest = String(data.get('interest') ?? '미정')
    const memo = String(data.get('memo') ?? '').trim()

    if (!name) return showMessage('이름을 입력해주세요.', 'error')
    if (!department) return showMessage('학과를 입력해주세요.', 'error')
    if (!grade) return showMessage('학년을 선택해주세요.', 'error')
    if (!email) return showMessage('이메일을 입력해주세요.', 'error')
    if (!activity) return showMessage('희망 활동을 선택해주세요.', 'error')
    if (selectedDays.length === 0) return showMessage('가능한 요일을 한 개 이상 선택해주세요.', 'error')
    if (selectedRoles.length === 0) return showMessage('희망 역할을 한 개 이상 선택해주세요.', 'error')
    if (startTime && endTime && startTime >= endTime) {
      return showMessage('종료 시간은 시작 시간보다 늦어야 합니다.', 'error')
    }

    const applicant: Applicant = {
      id: Date.now(),
      name,
      department,
      grade,
      email,
      phone,
      startDate,
      activity,
      days: selectedDays,
      startTime,
      endTime,
      roles: selectedRoles,
      interest,
      memo,
      createdAt: new Date().toLocaleString('ko-KR'),
    }

    setApplicants((prev) => [applicant, ...prev])
    showMessage(`${name}님의 신청 정보가 저장되었습니다.`)
    event.currentTarget.reset()
    nameRef.current?.focus()
  }

  const deleteApplicant = (id: number) => {
    const target = applicants.find((person) => person.id === id)
    if (!target) return
    if (!window.confirm(`${target.name}님의 신청 정보를 삭제할까요?`)) return
    setApplicants((prev) => prev.filter((person) => person.id !== id))
  }

  const clearAll = () => {
    if (applicants.length === 0) {
      window.alert('삭제할 데이터가 없습니다.')
      return
    }
    if (!window.confirm('저장된 모든 신청 정보를 삭제할까요? 이 작업은 되돌릴 수 없습니다.')) return
    setApplicants([])
  }

  const downloadExcel = () => {
    if (applicants.length === 0) {
      window.alert('엑셀로 저장할 신청 데이터가 없습니다.')
      return
    }

    const excelRows = applicants.map((person, index) => ({
      번호: index + 1,
      이름: person.name,
      학과: person.department,
      학년: person.grade,
      이메일: person.email,
      연락처: person.phone,
      희망활동: person.activity,
      가능한요일: person.days.join(', '),
      가능시작시간: person.startTime,
      가능종료시간: person.endTime,
      가능시작일: person.startDate,
      희망역할: person.roles.join(', '),
      관심분야: person.interest,
      소개및요청사항: person.memo,
      신청일시: person.createdAt,
    }))

    const worksheet = XLSX.utils.json_to_sheet(excelRows)
    worksheet['!cols'] = [
      { wch: 7 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 28 },
      { wch: 16 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 14 },
      { wch: 14 }, { wch: 24 }, { wch: 18 }, { wch: 34 }, { wch: 22 },
    ]

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, '신청자목록')

    const now = new Date()
    const dateString = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`
    XLSX.writeFile(workbook, `Campus_Mate_신청자_${dateString}.xlsx`)
  }

  return (
    <>
      <header className="hero">
        <div className="hero-inner">
          <div className="brand-mark" aria-hidden="true">🎓</div>
          <div>
            <p className="eyebrow">TEAM & STUDY MATCHING</p>
            <h1>Campus Mate</h1>
            <p className="subtitle">팀플과 스터디, 조건이 맞는 사람을 더 빠르게 만나보세요.</p>
          </div>
        </div>
      </header>

      <main className="page-shell">
        <section className="stats-grid" aria-label="신청 현황 요약">
          <article className="stat-card"><span>전체 신청</span><strong>{stats.total}</strong><small>명</small></article>
          <article className="stat-card"><span>팀플 희망</span><strong>{stats.team}</strong><small>명</small></article>
          <article className="stat-card"><span>스터디 희망</span><strong>{stats.study}</strong><small>명</small></article>
        </section>

        <section className="layout-grid">
          <article className="panel form-panel">
            <div className="panel-heading">
              <div><p className="section-kicker">APPLICATION</p><h2>신청 정보 입력</h2></div>
              <span className="badge">필수 항목 *</span>
            </div>

            <form ref={formRef} onSubmit={handleSubmit} onReset={() => setMessage(null)} noValidate>
              <div className="form-grid two-col">
                <div className="field-group"><label htmlFor="name">이름 <span>*</span></label><input ref={nameRef} type="text" id="name" name="name" placeholder="예: 홍길동" autoComplete="name" required /></div>
                <div className="field-group"><label htmlFor="department">학과 <span>*</span></label><input type="text" id="department" name="department" placeholder="예: 컴퓨터공학과" required /></div>
                <div className="field-group"><label htmlFor="grade">학년 <span>*</span></label><select id="grade" name="grade" required defaultValue=""><option value="">선택하세요</option><option>1학년</option><option>2학년</option><option>3학년</option><option>4학년 이상</option></select></div>
                <div className="field-group"><label htmlFor="email">이메일 <span>*</span></label><input type="email" id="email" name="email" placeholder="example@univ.ac.kr" autoComplete="email" required /></div>
                <div className="field-group"><label htmlFor="phone">연락처</label><input type="tel" id="phone" name="phone" placeholder="010-1234-5678" autoComplete="tel" /></div>
                <div className="field-group"><label htmlFor="startDate">가능 시작일</label><input type="date" id="startDate" name="startDate" /></div>
              </div>

              <fieldset className="choice-block">
                <legend>희망 활동 <span>*</span></legend>
                <div className="choice-row">
                  <label className="choice-chip"><input type="radio" name="activity" value="팀플" /><span>👥 팀플</span></label>
                  <label className="choice-chip"><input type="radio" name="activity" value="스터디" /><span>📚 스터디</span></label>
                  <label className="choice-chip"><input type="radio" name="activity" value="둘 다 가능" /><span>✨ 둘 다 가능</span></label>
                </div>
              </fieldset>

              <fieldset className="choice-block">
                <legend>가능한 요일 <span>*</span></legend>
                <div className="choice-row wrap">{days.map((day) => <label className="choice-chip" key={day}><input type="checkbox" name="days" value={day} /><span>{day}</span></label>)}</div>
              </fieldset>

              <div className="form-grid two-col">
                <div className="field-group"><label htmlFor="startTime">가능 시작 시간</label><input type="time" id="startTime" name="startTime" /></div>
                <div className="field-group"><label htmlFor="endTime">가능 종료 시간</label><input type="time" id="endTime" name="endTime" /></div>
              </div>

              <fieldset className="choice-block">
                <legend>희망 역할 <span>*</span></legend>
                <div className="choice-row wrap">{roles.map((role) => <label className="choice-chip" key={role}><input type="checkbox" name="roles" value={role} /><span>{role}</span></label>)}</div>
              </fieldset>

              <div className="field-group"><label htmlFor="interest">관심 분야</label><select id="interest" name="interest" defaultValue="미정"><option value="미정">선택하세요</option><option>IT·개발</option><option>공학·기술</option><option>경영·경제</option><option>디자인·콘텐츠</option><option>어학·자격증</option><option>취업·면접</option><option>기타</option></select></div>
              <div className="field-group"><label htmlFor="memo">한 줄 소개 / 요청사항</label><textarea id="memo" name="memo" rows={4} placeholder="예: 주 2회 저녁 시간에 꾸준히 참여할 수 있습니다." /></div>

              {message && <div className={`form-message show ${message.type}`} role="status" aria-live="polite">{message.text}</div>}

              <div className="button-row"><button type="submit" className="btn btn-primary">신청 정보 저장</button><button type="reset" className="btn btn-secondary">다시 작성</button></div>
            </form>
          </article>

          <aside className="panel guide-panel">
            <div className="guide-icon">💡</div><h2>사용 방법</h2>
            <ol><li>신청 정보를 입력합니다.</li><li><strong>신청 정보 저장</strong>을 누르면 브라우저에 자동 저장됩니다.</li><li>아래 신청자 목록에서 검색하거나 조건별로 필터링할 수 있습니다.</li><li><strong>Excel 다운로드</strong>를 누르면 현재 저장된 전체 자료가 <code>.xlsx</code>로 내려받아집니다.</li></ol>
            <div className="notice-box"><strong>자료 저장 안내</strong><p>이 버전은 브라우저의 localStorage를 사용하므로 새로고침해도 자료가 유지됩니다.</p></div>
          </aside>
        </section>

        <section className="panel list-panel">
          <div className="panel-heading responsive-heading">
            <div><p className="section-kicker">APPLICANTS</p><h2>신청자 관리</h2><p className="muted">저장된 신청자를 검색·필터링하고 엑셀로 내보낼 수 있습니다.</p></div>
            <div className="export-actions"><button type="button" onClick={downloadExcel} className="btn btn-excel">📊 Excel 다운로드</button><button type="button" onClick={clearAll} className="btn btn-danger-outline">전체 삭제</button></div>
          </div>

          <div className="toolbar">
            <div className="search-box"><label htmlFor="searchInput" className="sr-only">신청자 검색</label><span aria-hidden="true">🔎</span><input type="search" id="searchInput" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="이름, 학과, 역할, 관심 분야 검색" /></div>
            <div className="filter-box"><label htmlFor="activityFilter">활동</label><select id="activityFilter" value={activityFilter} onChange={(e) => setActivityFilter(e.target.value)}><option value="all">전체</option><option>팀플</option><option>스터디</option><option>둘 다 가능</option></select></div>
            <div className="filter-box"><label htmlFor="dayFilter">요일</label><select id="dayFilter" value={dayFilter} onChange={(e) => setDayFilter(e.target.value)}><option value="all">전체</option>{days.map((day) => <option key={day}>{day}</option>)}</select></div>
          </div>

          <div className="table-wrap">
            <table><thead><tr><th>번호</th><th>이름</th><th>학과 / 학년</th><th>희망 활동</th><th>가능 요일</th><th>희망 역할</th><th>시간</th><th>관심 분야</th><th>관리</th></tr></thead>
              <tbody>
                {filteredApplicants.map((person, index) => {
                  const timeText = person.startTime || person.endTime ? `${person.startTime || '-'} ~ ${person.endTime || '-'}` : '미정'
                  return <tr key={person.id}>
                    <td>{index + 1}</td>
                    <td><div className="cell-title">{person.name}</div><div className="cell-sub">{person.email}</div></td>
                    <td><div className="cell-title">{person.department}</div><div className="cell-sub">{person.grade}</div></td>
                    <td><span className="mini-tag">{person.activity}</span></td>
                    <td><div className="tag-list">{person.days.map((day) => <span className="mini-tag" key={day}>{day}</span>)}</div></td>
                    <td><div className="tag-list">{person.roles.map((role) => <span className="mini-tag" key={role}>{role}</span>)}</div></td>
                    <td><div>{timeText}</div><div className="cell-sub">{person.startDate ? `시작 ${person.startDate}` : '시작일 미정'}</div></td>
                    <td>{person.interest}</td>
                    <td><button className="delete-btn" type="button" onClick={() => deleteApplicant(person.id)}>삭제</button></td>
                  </tr>
                })}
              </tbody>
            </table>
            {filteredApplicants.length === 0 && <div className="empty-state show"><div>🗂️</div><strong>조건에 맞는 신청자가 없습니다.</strong><p>신청서를 작성하거나 검색 조건을 바꿔보세요.</p></div>}
          </div>
        </section>
      </main>

      <footer><p>Campus Mate · 대학생 팀플/스터디 신청 관리</p></footer>
    </>
  )
}

export default App
