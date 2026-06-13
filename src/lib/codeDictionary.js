/* Extracted from the validated single-file build — logic unchanged. */
const DEFAULT_CODES = {
  // standalone
  P:{lop:0,el:0,cl:0,ml:0,co:0,eco:0,present:1,half:0,kind:'present'},
  L:{lop:1,present:0,kind:'lop'},
  WO:{kind:'weekoff'}, HO:{kind:'holiday'}, NH:{kind:'holiday'}, FH:{kind:'holiday'},
  EL:{el:1,kind:'leave'}, CL:{cl:1,kind:'leave'}, ML:{ml:1,kind:'leave'}, OD:{present:1,kind:'present'},
  MSP:{kind:'missing'}, LEFT:{kind:'status'},
  // half / combined  (first half / second half)
  'P/A':{lop:.5,present:.5,half:1,kind:'half'}, 'A/P':{lop:.5,present:.5,half:1,kind:'half'},
  'EL/P':{el:.5,present:.5,half:1,kind:'half'}, 'P/EL':{el:.5,present:.5,half:1,kind:'half'},
  'EL/A':{el:.5,lop:.5,half:1,kind:'half'}, 'A/EL':{el:.5,lop:.5,half:1,kind:'half'},
  'CL/P':{cl:.5,present:.5,half:1,kind:'half'}, 'P/CL':{cl:.5,present:.5,half:1,kind:'half'},
  'CL/A':{cl:.5,lop:.5,half:1,kind:'half'}, 'A/CL':{cl:.5,lop:.5,half:1,kind:'half'},
  'EL/L':{el:.5,lop:.5,half:1,kind:'half'}, 'L/EL':{el:.5,lop:.5,half:1,kind:'half'},
  'P/L':{present:.5,lop:.5,half:1,kind:'half'}, 'L/P':{present:.5,lop:.5,half:1,kind:'half'},
  'CL/L':{cl:.5,lop:.5,half:1,kind:'half'}, 'L/CL':{cl:.5,lop:.5,half:1,kind:'half'},
  // comp-off (supported, never auto-applied)
  CO:{co:1,kind:'compoff'}, 'CO/2':{co:.5,kind:'compoff'},
  ECO:{eco:1,kind:'compoff'}, 'ECO/2':{eco:.5,kind:'compoff'},
  'P/CO/2':{present:.5,co:.5,half:1,kind:'compoff'}, 'CO/2/P':{present:.5,co:.5,half:1,kind:'compoff'},
  'CO/EL':{co:.5,el:.5,half:1,kind:'compoff'},
};
const codeMeta = (code, dict) => dict[String(code||'').trim()] || null;

export { DEFAULT_CODES, codeMeta };
