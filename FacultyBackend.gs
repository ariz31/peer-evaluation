function unlockFaculty(pin){ensureSetup();if(String(pin||'')!==pin_())return fail_('Incorrect faculty code.');return ok_('Unlocked.',facultyState_())}
