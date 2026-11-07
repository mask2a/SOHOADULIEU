"""Pure business decisions; no Tkinter or Excel imports."""
from .domain import MatchResult, MATCH, NORMALIZED, REVIEW, ONLY_SOURCE, ONLY_SCHOOL, IDENTITY

def reconcile(store, check=lambda: None):
    # Complete exact matches first, so approximate candidates cannot consume them.
    completed_single_source_keys = set()
    for school_id, school in store.query("side='school' AND valid=1"):
        check()
        sources = store.exact(school, 'source')
        schools = store.exact(school, 'school')
        if len(sources) == 1 and schools:
            if school.keys in completed_single_source_keys:
                continue
            completed_single_source_keys.add(school.keys)
            source_id, source = sources[0]
            store.mark([source_id, *(record_id for record_id, _ in schools)])
            matching_schools = [record for _, record in schools]
            raw_equal = all(source.values[i] == matching_schools[0].values[i] for i in IDENTITY)
            locations = '; '.join(record.location for record in matching_schools)
            reason = (f'Có {len(matching_schools)} hồ sơ file trường khớp thông tin: '
                      f'{locations}')
            yield MatchResult(
                MATCH if raw_equal else NORMALIZED,
                source,
                matching_schools[0],
                reason,
                matching_schools=matching_schools,
            )
            continue
        pairs = []
        if len(sources) == len(schools) == 1:
            pairs = sources
        elif school.certificate:
            source_cert = [(i,r) for i,r in sources if r.certificate == school.certificate]
            school_cert = [(i,r) for i,r in schools if r.certificate == school.certificate]
            if len(source_cert) == len(school_cert) == 1:
                pairs = source_cert
        if pairs:
            source_id, source = pairs[0]
            store.mark([source_id, school_id])
            raw_equal = all(source.values[i] == school.values[i] for i in IDENTITY)
            yield MatchResult(MATCH if raw_equal else NORMALIZED, source, school,
                '' if raw_equal else 'Bộ ba bằng nhau sau chuẩn hóa; xem giá trị gốc.')

    source_candidates = {}
    for school_id, school in store.query("side='school' AND used=0"):
        check()
        candidates = store.candidates(school)
        store.mark([school_id])
        if candidates:
            for source_id, source in candidates:
                source_candidates.setdefault(source_id, []).append(school)
            reason = 'Chưa xác định được cặp; không tính khớp.'
            if any(any(not r.values[k].strip() and school.values[k].strip() for k in IDENTITY) for _,r in candidates):
                reason += ' Có thông tin trường đề nghị bổ sung; phải xác minh đúng người trước.'
            if any(any(r.values[k].strip() and not school.values[k].strip() for k in IDENTITY) for _,r in candidates):
                reason += ' Trường thiếu thông tin so với ứng viên nguồn chuẩn.'
            yield MatchResult(REVIEW, school=school, reason=reason, candidates=[r for _,r in candidates])
        elif not school.valid:
            reason = 'Thiếu/sai họ tên, ngày sinh hoặc năm tốt nghiệp; ngày sinh chỉ năm chưa đủ.'
            yield MatchResult(REVIEW, school=school, reason=reason)
        else:
            yield MatchResult(ONLY_SCHOOL, school=school)

    for source_id, source in store.query("side='source' AND used=0"):
        check()
        store.mark([source_id])
        candidates = source_candidates.get(source_id, [])
        if candidates:
            reason = f'Chưa ghép chính xác; có {len(candidates)} hồ sơ trường cần xác minh.'
            status = REVIEW
        elif not source.valid:
            reason = 'Nguồn chuẩn thiếu/sai định danh hoặc ngày sinh chưa đầy đủ.'
            status = REVIEW
        else:
            reason = ''
            status = ONLY_SOURCE
        yield MatchResult(status, source=source, reason=reason, candidates=candidates)
