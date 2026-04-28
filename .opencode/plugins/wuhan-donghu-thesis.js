/**
 * Wuhan Donghu Thesis plugin for OpenCode.
 *
 * Registers this repository's bundled skills directory without requiring
 * symlinks or manual skills.paths edits.
 */

import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export const WuhanDonghuThesisPlugin = async () => {
  const skillsDir = path.resolve(__dirname, '../../skills');

  return {
    config: async (config) => {
      config.skills = config.skills || {};
      config.skills.paths = config.skills.paths || [];
      if (!config.skills.paths.includes(skillsDir)) {
        config.skills.paths.push(skillsDir);
      }
    }
  };
};
