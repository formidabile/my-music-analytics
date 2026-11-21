const content = document.getElementById('page-content');
const page = document.body.dataset.page;
const chartInstances = new Map();

async function getJSON(url) {
    const response = await fetch(url);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || 'Не удалось загрузить данные');
    return payload;
}

function number(value) {
    return new Intl.NumberFormat('ru-RU').format(value ?? 0);
}

function empty(message) {
    content.innerHTML = `<div class="empty-state">${message}</div>`;
}

function table(headers, rows) {
    return `<div class="panel"><div class="panel-header"><h2 class="panel-title">${headers.title}</h2></div><table class="data-table"><thead><tr>${headers.columns.map(x => `<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${row.map(cell => `<td>${cell ?? '—'}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}

async function renderOverview() {
    const data = await getJSON('/api/overview');
    if (!data.database_ready) {
        empty('База данных пока не подготовлена. Сначала запусти импорт рейтингов, прогнозов и стримингов — после этого здесь появится аналитика.');
        return;
    }
    const latestRows = data.latest_top.map(x => [x.position, `<a href="/songs/${x.song_id}"><strong>${escapeHTML(x.title)}</strong></a>`, escapeHTML(x.artists || '—'), x.points]);
    const artistRows = data.top_artists.map((x, i) => [i + 1, `<a href="/artists/${x.artist_id}">${escapeHTML(x.artist)}</a>`, x.appearances, number(x.total_points)]);
    const streamingRows = data.streaming_top.map(x => [x.position, escapeHTML(x.artist), x.listening_minutes == null ? 'Нет данных' : number(x.listening_minutes)]);
    const metrics = data.forecast_metrics;
    const precision = metrics.length ? `${(metrics[metrics.length - 1].precision * 100).toFixed(1).replace('.', ',')}%` : '—';
    content.innerHTML = `
      <div class="grid kpi-grid">
        ${kpi('Песен в базе', number(data.songs_count), 'Уникальные композиции')}
        ${kpi('Исполнителей', number(data.artists_count), 'По фактическим рейтингам')}
        ${kpi('Лет рейтингов', number(data.years.length), data.years.length ? `${data.years[0]}–${data.years.at(-1)}` : 'Нет данных')}
        ${kpi('Последний рейтинг', data.latest_year ?? '—', 'Год последнего Top-50')}
      </div>
      <div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Динамика рейтингов</h2></div><div class="chart-wrap"><canvas id="overview-years-chart"></canvas></div></div>
      <div class="grid two-column">
        ${table({title:`Top-5 · ${data.latest_year ?? ''}`,columns:['#','Песня','Исполнитель','Баллы']},latestRows)}
        ${table({title:'Топ исполнителей за всё время',columns:['#','Исполнитель','Попадания','Баллы']},artistRows)}
        ${table({title:'Стриминговый Top-5',columns:['#','Исполнитель','Минуты']},streamingRows)}
        <div class="panel"><div class="panel-header"><div><h2 class="panel-title">Точность прогнозов</h2><div class="panel-subtitle">Precision последнего года с прогнозом</div></div><span class="pill">${metrics.length ? metrics[metrics.length - 1].year : 'Нет данных'}</span></div><div class="kpi-value">${precision}</div><div class="kpi-note">Детальные показатели доступны в разделе «Прогнозы»</div></div>
      </div>`;
    drawChart('overview-years-chart', {type:'bar', data:{labels:data.yearly.map(x=>x.year), datasets:[{label:'Новинки',data:data.yearly.map(x=>x.new_songs)},{label:'Повторные попадания',data:data.yearly.map(x=>x.repeated_songs)},{label:'Возвращения',data:data.yearly.map(x=>x.returning_songs)}]}, options:chartOptions()});
    if (metrics.length) {
        document.querySelector('.chart-panel').insertAdjacentHTML('afterend', '<div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Точность прогнозов по годам</h2></div><div class="chart-wrap"><canvas id="overview-forecast-chart"></canvas></div></div>');
        drawChart('overview-forecast-chart', {type:'line', data:{labels:metrics.map(x=>x.year), datasets:[{label:'Precision',data:metrics.map(x=>+(x.precision*100).toFixed(1)),tension:.25},{label:'Recall',data:metrics.map(x=>+(x.recall*100).toFixed(1)),tension:.25}]}, options:{...chartOptions(),scales:{...chartOptions().scales,y:{...chartOptions().scales.y,min:0,max:100,ticks:{color:'#a9bad1',callback:value=>value+'%'}}}}});
    }
}

function kpi(label, value, note) {
    return `<div class="kpi-card"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div><div class="kpi-note">${note}</div></div>`;
}

function escapeHTML(value) {
    return String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
}

async function renderRankings() {
    const {years} = await getJSON('/api/years');
    if (!years.length) return empty('Пока нет импортированных рейтингов.');
    content.innerHTML = `<div class="toolbar"><label for="year-select" class="muted">Год рейтинга</label><select id="year-select">${years.slice().reverse().map(y => `<option value="${y}">${y}</option>`).join('')}</select><input id="ranking-search" placeholder="Поиск по песне или исполнителю"></div><div id="ranking-composition" class="grid two-column"><div id="ranking-age"></div><div id="ranking-countries"></div><div id="ranking-languages"></div></div><div id="ranking-table"></div>`;
    const select = document.getElementById('year-select');
    const search = document.getElementById('ranking-search');
    let selectedRanking = null;

    function renderComposition(composition, year) {
        const age = composition.age || [];
        const countries = composition.countries || [];
        const languages = composition.languages || [];
        document.getElementById('ranking-age').innerHTML = `
            <div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Возрастной состав · ${year}</h2></div><div class="chart-wrap chart-wrap-small"><canvas id="ranking-age-chart"></canvas></div></div>
            ${table({title:'Возрастные категории',columns:['Категория','Песен']},age.map(x=>[escapeHTML(x.category),x.songs_count]))}`;
        document.getElementById('ranking-countries').innerHTML = `
            <div class="panel chart-panel"><div class="panel-header"><div><h2 class="panel-title">Страны · доля песен</h2><div class="panel-subtitle">Одна песня может относиться к нескольким странам</div></div></div><div class="chart-wrap chart-wrap-small"><canvas id="ranking-countries-chart"></canvas></div></div>
            ${table({title:'Распределение по странам',columns:['Страна','Песен']},countries.map(x=>[escapeHTML(x.category),x.songs_count]))}`;
        document.getElementById('ranking-languages').innerHTML = `
            <div class="panel chart-panel"><div class="panel-header"><div><h2 class="panel-title">Языки · доля песен</h2><div class="panel-subtitle">Одна песня может относиться к нескольким языкам</div></div></div><div class="chart-wrap chart-wrap-small"><canvas id="ranking-languages-chart"></canvas></div></div>
            ${table({title:'Распределение по языкам',columns:['Язык','Песен']},languages.map(x=>[escapeHTML(x.category),x.songs_count]))}`;

        drawChart('ranking-age-chart', {type:'doughnut', data:{labels:age.map(x=>x.category),datasets:[{label:'Доля песен',data:age.map(x=>+(x.share*100).toFixed(1))}]},options:{...chartOptions(),plugins:{legend:{position:'right',labels:{color:'#a9bad1'}}}}});
        const shareOptions = {...chartOptions(),indexAxis:'y',scales:{...chartOptions().scales,x:{...chartOptions().scales.x,min:0,max:100,ticks:{color:'#a9bad1',callback:value=>value+'%'}}}};
        const countryChart = countries.slice(0,10);
        drawChart('ranking-countries-chart', {type:'bar',data:{labels:countryChart.map(x=>x.category),datasets:[{label:'Доля песен, %',data:countryChart.map(x=>+(x.share*100).toFixed(1))}]},options:shareOptions});
        const languageChart = languages.slice(0,10);
        drawChart('ranking-languages-chart', {type:'bar',data:{labels:languageChart.map(x=>x.category),datasets:[{label:'Доля песен, %',data:languageChart.map(x=>+(x.share*100).toFixed(1))}]},options:shareOptions});
    }

    function renderRankingTable() {
        if (!selectedRanking) return;
        const query = search.value.trim().toLocaleLowerCase('ru');
        const rows = selectedRanking.entries.filter(x => `${x.title} ${x.artists}`.toLocaleLowerCase('ru').includes(query));
        document.getElementById('ranking-table').innerHTML = table({title:`Top-50 · ${selectedRanking.year}`,columns:['#','Песня','Исполнитель','Год выпуска','Баллы']},rows.map(x => [x.position,`<a href="/songs/${x.song_id}">${escapeHTML(x.title)}</a>`,escapeHTML(x.artists),x.release_year,x.points]));
    }

    async function loadYear() {
        selectedRanking = await getJSON(`/api/rankings/${select.value}`);
        renderComposition(selectedRanking.composition, selectedRanking.year);
        renderRankingTable();
    }
    select.addEventListener('change', loadYear);
    search.addEventListener('input', renderRankingTable);
    await loadYear();
}

async function renderAnalytics() {
    const data = await getJSON('/api/analytics');
    content.innerHTML = `<div class="grid kpi-grid">${kpi('Песен',number(data.songs.length),'Вся история')}${kpi('Исполнителей',number(data.artists.length),'Вся история')}${kpi('Hall of Fame',number(data.hall_of_fame.length),'Порог: 3 появления')}${kpi('Анализ исполнителей',number(data.stability.length),'Показатели устойчивости')}</div><div class="grid two-column"><div id="artists-panel"></div><div id="hof-panel"></div><div id="yearly-panel"></div><div id="composition-panel"></div><div id="stability-panel"></div></div><div id="song-stability-panel"></div>`;
    document.getElementById('artists-panel').innerHTML = table({title:'Топ исполнителей',columns:['#','Исполнитель','Попадания','Баллы']},data.artists.slice(0,15).map((x,i)=>[i+1,`<a href="/artists/${x.artist_id}">${escapeHTML(x.artist)}</a>`,x.appearances,number(x.total_points)]));
    document.getElementById('hof-panel').innerHTML = table({title:'Hall of Fame',columns:['Песня','Исполнитель','Появления','Год включения']},data.hall_of_fame.slice(0,15).map(x=>[escapeHTML(x.title),escapeHTML(x.artists.join('; ')),x.appearances,x.induction_year]));
    document.getElementById('yearly-panel').innerHTML = table({title:'Динамика рейтингов',columns:['Год','Песен','Новинки','Повторные попадания','Возвращения']},data.yearly.map(x=>[x.year,x.entries,x.new_songs,x.repeated_songs,x.returning_songs]));
    const composition = data.composition.filter(x=>x.year === null && x.dimension === 'Страна').sort((a,b)=>b.songs_count-a.songs_count).slice(0,12);
    document.getElementById('composition-panel').innerHTML = table({title:'Страны · вся история',columns:['Страна','Песен','Доля']},composition.map(x=>[escapeHTML(x.category),x.songs_count,`${(x.share*100).toFixed(1)}%`]));
    const stableSongs = data.song_stability;
    document.getElementById('song-stability-panel').innerHTML = table({title:'Устойчивость песен',columns:['Песня','Исполнитель','Лет в рейтинге','Доля лет','Самая длинная серия','Текущая серия']},stableSongs.map(x=>[
        `<a href="/songs/${x.song_id}">${escapeHTML(x.title)}</a>`, escapeHTML(x.artists.join('; ')),
        `${x.years_present} из ${x.total_ranking_years}`, `${(x.presence_rate*100).toFixed(1)}%`,
        x.longest_streak, x.current_streak
    ]));
    const stableArtists = data.stability.slice(0,15);
    document.getElementById('stability-panel').innerHTML = table({title:'Устойчивость исполнителей',columns:['Исполнитель','Лет в рейтинге','Доля лет','Самая длинная серия','Песен за активный год']},stableArtists.map(x=>[escapeHTML(x.artist),`${x.years_present} из ${x.total_ranking_years}`,`${(x.presence_rate*100).toFixed(1)}%`,x.longest_streak,Number(x.average_tracks_per_present_year).toFixed(1)]));
    document.getElementById('stability-panel').insertAdjacentHTML('beforebegin', '<div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Регулярность присутствия исполнителей</h2></div><div class="chart-wrap"><canvas id="artist-presence-chart"></canvas></div></div>');
    drawChart('artist-presence-chart', {type:'bar', data:{labels:stableArtists.slice(0,10).map(x=>x.artist),datasets:[{label:'Доля лет в рейтинге, %',data:stableArtists.slice(0,10).map(x=>+(x.presence_rate*100).toFixed(1))}]},options:{...chartOptions(),indexAxis:'y',scales:{...chartOptions().scales,x:{...chartOptions().scales.x,min:0,max:100,ticks:{color:'#a9bad1',callback:value=>value+'%'}}}}});
    document.getElementById('yearly-panel').insertAdjacentHTML('beforebegin', '<div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Состав рейтинга по годам</h2></div><div class="chart-wrap"><canvas id="analytics-years-chart"></canvas></div></div>');
    drawChart('analytics-years-chart', {type:'bar', data:{labels:data.yearly.map(x=>x.year), datasets:[{label:'Новинки',data:data.yearly.map(x=>x.new_songs)},{label:'Повторные попадания',data:data.yearly.map(x=>x.repeated_songs)},{label:'Возвращения',data:data.yearly.map(x=>x.returning_songs)}]}, options:chartOptions()});
    const years = data.yearly.map(x=>x.year);
    document.getElementById('artists-panel').insertAdjacentHTML('afterend', '<div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Баллы ведущих исполнителей по годам</h2></div><div class="chart-wrap"><canvas id="artist-points-chart"></canvas></div></div>');
    const artistColors = ['#63a9ef','#53c7a0','#e7b66b','#b69af5','#e889a5','#8da7c7'];
    drawChart('artist-points-chart', {type:'line', data:{labels:years, datasets:data.top_artists_yearly.map((artist,i)=>({label:artist.artist,data:years.map(year=>artist.points_by_year[year] || 0),borderColor:artistColors[i],backgroundColor:artistColors[i],tension:.25}))}, options:chartOptions()});
    const languages = data.composition.filter(x=>x.year === null && x.dimension === 'Язык').sort((a,b)=>b.songs_count-a.songs_count).slice(0,8);
    const decades = data.composition.filter(x=>x.year === null && x.dimension === 'Десятилетие').sort((a,b)=>a.category.localeCompare(b.category,'ru'));
    document.getElementById('composition-panel').insertAdjacentHTML('beforeend', '<div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Песни по странам</h2></div><div class="chart-wrap chart-wrap-small"><canvas id="countries-chart"></canvas></div></div><div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Песни по языкам</h2></div><div class="chart-wrap chart-wrap-small"><canvas id="languages-chart"></canvas></div></div><div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Песни по десятилетиям выпуска</h2></div><div class="chart-wrap chart-wrap-small"><canvas id="decades-chart"></canvas></div></div>');
    drawChart('countries-chart', {type:'bar', data:{labels:composition.slice(0,8).map(x=>x.category), datasets:[{label:'Песен',data:composition.slice(0,8).map(x=>x.songs_count)}]}, options:{...chartOptions(),indexAxis:'y'}});
    drawChart('languages-chart', {type:'doughnut', data:{labels:languages.map(x=>x.category), datasets:[{label:'Песен',data:languages.map(x=>x.songs_count)}]}, options:{...chartOptions(),plugins:{legend:{position:'right',labels:{color:'#a9bad1'}}}}});
    drawChart('decades-chart', {type:'bar', data:{labels:decades.map(x=>x.category), datasets:[{label:'Песен',data:decades.map(x=>x.songs_count)}]}, options:chartOptions()});
}

async function renderStreaming() {
    const data = await getJSON('/api/streaming');
    if (!data.entries.length) return empty('Стриминговые данные пока не импортированы.');
    content.innerHTML = `<div class="toolbar"><label for="stream-year-select" class="muted">Год рейтинга</label><select id="stream-year-select">${data.years.slice().reverse().map(y => `<option value="${y}">${y}</option>`).join('')}</select></div><div id="stream-year-table"></div><div class="panel chart-panel"><div class="panel-header"><h2 class="panel-title">Время прослушивания · выбранный год</h2></div><div class="chart-wrap"><canvas id="streaming-chart"></canvas></div></div><div id="stream-all-table"></div>`;
    const select = document.getElementById('stream-year-select');
    function loadYear() {
        const entries = data.entries.filter(x => Number(x.year) === Number(select.value));
        document.getElementById('stream-year-table').innerHTML = table({title:`Стриминговый Top-5 · ${select.value}`,columns:['#','Исполнитель','Минуты']},entries.map(x=>[x.position,escapeHTML(x.artist),x.listening_minutes == null ? 'Нет данных' : number(x.listening_minutes)]));
        const known = entries.filter(x=>x.listening_minutes != null);
        drawChart('streaming-chart', {type:'bar', data:{labels:known.map(x=>x.artist),datasets:[{label:'Минуты прослушивания',data:known.map(x=>x.listening_minutes)}]},options:{...chartOptions(),indexAxis:'y'}});
    }
    select.addEventListener('change', loadYear);
    loadYear();
    document.getElementById('stream-all-table').innerHTML = table({title:'За всё время · сортировка по минутам',columns:['Исполнитель','Лет в Top-5','Минуты']},data.all_time.map(x=>[escapeHTML(x.artist),x.years_count,x.known_minutes_years ? number(x.total_minutes) : 'Нет данных']));
}

function chartOptions() {
    return {responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#a9bad1'}}},scales:{x:{ticks:{color:'#a9bad1'},grid:{color:'rgba(148,163,184,.10)'}},y:{ticks:{color:'#a9bad1'},grid:{color:'rgba(148,163,184,.10)'}}}};
}

function drawChart(id, config) {
    const canvas = document.getElementById(id);
    if (!canvas || typeof Chart === 'undefined') return;
    if (chartInstances.has(id)) chartInstances.get(id).destroy();
    chartInstances.set(id, new Chart(canvas, config));
}

async function renderForecasts() {
    const data = await getJSON('/api/forecasts');
    if (!data.metrics.length) return empty('Данные прогнозов пока не импортированы.');
    const metricRows = data.metrics.map(x=>[x.year,x.predicted_count,x.hits,x.unexpected,`${(x.precision*100).toFixed(1)}%`,`${(x.recall*100).toFixed(1)}%`,x.mae == null ? '—' : x.mae.toFixed(2)]);
    const unexpectedRows = data.unexpected.slice(0,30).map(x=>[x.forecast_year,escapeHTML(x.title),escapeHTML(x.artists.join('; ')),x.actual_position]);
    const delayedRows = data.delayed.map(x=>[x.forecast_year,escapeHTML(x.title),escapeHTML(x.artists.join('; ')),x.first_actual_year,x.delay_years]);
    const pendingRows = data.comparisons.filter(x=>x.status==='PENDING').slice(0,30).map(x=>[x.forecast_year,x.predicted_position,escapeHTML(x.title),escapeHTML(x.artists.join('; '))]);
    content.innerHTML = `<div class="grid">${table({title:'Эффективность прогнозов',columns:['Год','Прогнозов','Совпадений','Неожиданных','Precision','Recall','MAE']},metricRows)}<div class="grid two-column"><div id="unexpected"></div><div id="delayed"></div><div id="pending"></div></div></div>`;
    document.getElementById('unexpected').innerHTML = table({title:'Неожиданные хиты',columns:['Год','Песня','Исполнитель','Позиция']},unexpectedRows);
    document.getElementById('delayed').innerHTML = table({title:'Отложенные хиты',columns:['Прогноз','Песня','Исполнитель','Первое попадание','Задержка']},delayedRows);
    document.getElementById('pending').innerHTML = table({title:'Нереализованные прогнозы',columns:['Год','Прогнозная позиция','Песня','Исполнитель']},pendingRows);
}

const renderers = {overview:renderOverview, rankings:renderRankings, analytics:renderAnalytics, streaming:renderStreaming, forecasts:renderForecasts};
(renderers[page] || renderOverview)().catch(error => empty(`Не удалось загрузить раздел: ${escapeHTML(error.message)}`));
