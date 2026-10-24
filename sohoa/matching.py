"""Pure business decisions; no Tkinter or Excel imports."""
from .domain import MatchResult, MATCH, NORMALIZED, REVIEW, ONLY_SOURCE, ONLY_SCHOOL, IDENTITY

def reconcile(store, check=lambda: None):
    # Complete exact matches first, so approximate candidates cannot consume them.
    for school_id, school in store.query("side='school' AND valid=1"):
        check()
        sources = store.exact(school, 'source')
        schools = store.exact(school, 'school')
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
    # Deferred rows remain review even if a candidate has already been matched elsewhere.
    for school_id, school in store.query("side='school' AND used=0"):
        check()
        candidates = store.candidates(school)
        if candidates or not school.valid:
            store.mark([school_id] + [i for i,_ in candidates])
            reason = 'Chưa xác định được cặp; không tính khớp.'
            if not school.valid:
                reason += ' Thiếu/sai họ tên, ngày sinh hoặc năm tốt nghiệp; ngày sinh chỉ năm chưa đủ.'
            if any(any(not r.values[k].strip() and school.values[k].strip() for k in IDENTITY) for _,r in candidates):
                reason += ' Có thông tin trường đề nghị bổ sung; phải xác minh đúng người trước.'
            if any(any(r.values[k].strip() and not school.values[k].strip() for k in IDENTITY) for _,r in candidates):
                reason += ' Trường thiếu thông tin so với ứng viên nguồn chuẩn.'
            yield MatchResult(REVIEW, school=school, reason=reason, candidates=[r for _,r in candidates])
        else:
            store.mark([school_id])
            yield MatchResult(ONLY_SCHOOL, school=school)
    for source_id, source in store.query("side='source' AND used=0"):
        check()
        store.mark([source_id])
        yield MatchResult(ONLY_SOURCE if source.valid else REVIEW, source=source,
            reason='' if source.valid else 'Nguồn chuẩn thiếu/sai định danh hoặc ngày sinh chưa đầy đủ.')
