// 入口函数不可修改，否则无法执行，args 为配置的入参 
export default async function main(args) { 
    // 1. 增加空值保护：如果 responseData 为空，默认设为 {} 
    let responseData = args.responseData || {}; 
    
    // 2. 提取所有课程信息：兼容 responseData.rows 为空的情况 
    let courseList = (responseData.rows && Array.isArray(responseData.rows)) 
        ? responseData.rows 
        : []; 
    
    // 生成当前日期和时间
    let currentDate = new Date(); 
    let formattedCurrentDate = currentDate.getFullYear() + '年' + 
        String(currentDate.getMonth() + 1).padStart(2, '0') + '月' + 
        String(currentDate.getDate()).padStart(2, '0') + '日';
    
    let generatedAt = currentDate.getFullYear() + '-' + 
        String(currentDate.getMonth() + 1).padStart(2, '0') + '-' + 
        String(currentDate.getDate()).padStart(2, '0') + ' ' + 
        String(currentDate.getHours()).padStart(2, '0') + ':' + 
        String(currentDate.getMinutes()).padStart(2, '0') + ':' + 
        String(currentDate.getSeconds()).padStart(2, '0');
    
    // 处理所有课程信息 
    let courses = []; 
    let teacherNameSet = new Set(); 
    
    for (let i = 0; i < courseList.length; i++) { 
        let courseInfo = courseList[i] || {}; 
        
        // 获取课程相关字段（增加空值保护） 
        let kcmc = courseInfo.kcmc || ''; 
        let kch = courseInfo.kch || ''; 
        let kkbm = courseInfo.kkbm || ''; 
        let bjmc = courseInfo.bjmc || ''; 
        let jsxm = courseInfo.jsxm || ''; 
        let xf = courseInfo.xf || ''; 
        let zxs = courseInfo.zxs || ''; 
        let kcxz = courseInfo.kcxz || ''; 
        let xnd = courseInfo.xnd || ''; 
        let xqm = courseInfo.xqm || ''; 
        
        // 收集教师姓名 
        if (jsxm) { 
            teacherNameSet.add(jsxm); 
        } 
        
        // 生成英文名称（简单处理） 
        let englishName = kcmc ? kcmc.replace(/[\u4e00-\u9fa5]/g, '') : ''; 
        if (!englishName) { 
            englishName = 'Course English Name'; 
        } 
        
        // 格式化学期 
        let formattedSemester = ''; 
        if (xnd && xqm) { 
            let yearMatch = xnd.match(/(\d{4})-(\d{4})/); 
            if (yearMatch) { 
                let startYear = parseInt(yearMatch[1]); 
                let endYear = parseInt(yearMatch[2]); 
                if (xqm === '1') { 
                    formattedSemester = `${startYear}学年秋季学期`; 
                } else if (xqm === '2') { 
                    formattedSemester = `${endYear}学年春季学期`; 
                } else { 
                    formattedSemester = `${xnd} 第${xqm}学期`; 
                } 
            } else { 
                formattedSemester = `${xnd} 第${xqm}学期`; 
            } 
        } 
        
        // 构建单门课程信息对象 
        let course = { 
            courseName: kcmc, 
            courseCode: kch, 
            department: kkbm, 
            applicableScope: bjmc, 
            teacherName: jsxm, 
            credits: xf, 
            totalHours: zxs, 
            courseNature: kcxz, 
            currentDate: formattedCurrentDate, 
            englishName: englishName, 
            formattedSemester: formattedSemester 
        }; 
        
        // 添加到courses数组 
        courses.push(course); 
    } 
    
    // 构建元数据 
    let metadata = { 
        generated_at: generatedAt, 
        total_courses: courses.length, 
        teacher_name: Array.from(teacherNameSet).join('、') || '', 
        formatted_date: formattedCurrentDate
    }; 
    
    // 返回结果 - 标准格式，包含metadata和courses
    return { 
        metadata: metadata, 
        courses: courses 
    }; 
}