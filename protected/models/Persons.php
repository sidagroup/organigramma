<?php

/**
 * This is the model class for table "persons".
 *
 * The followings are the available columns in table 'persons':
 * @property string $PersonID
 * @property string $Created
 * @property string $CreatedBy
 * @property integer $RoleID
 * @property string $UserName
 * @property string $FirstName
 * @property string $LastName
 * @property string $PersonEmail
 * @property string $Password
 * @property string $Gender
 * @property integer $Enabled
 * @property string $LastLogin
 *
 * The followings are the available model relations:
 * @property Roles $role
 * @property Persons $createdBy
 * @property Persons[] $persons
 * @property PersonsCities[] $personsCities
 * @property PersonsMasters[] $personsMasters
 */
class Persons extends CActiveRecord
{
	/**
	 * @return string the associated database table name
	 */
	public function tableName()
	{
		return 'persons';
	}

	/**
	 * @return array validation rules for model attributes.
	 */
	public function rules()
	{
		// NOTE: you should only define rules for those attributes that
		// will receive user inputs.
		return array(
			array('RoleID, UserName, FirstName, LastName, Password, Gender, Enabled', 'required'),
			array('RoleID, Enabled', 'numerical', 'integerOnly'=>true),
			array('CreatedBy', 'length', 'max'=>11),
			array('UserName, FirstName, LastName, PersonEmail', 'length', 'max'=>255),
			array('Password', 'length', 'max'=>32),
			array('Gender', 'length', 'max'=>1),
			array('Created, LastLogin', 'safe'),
			// The following rule is used by search().
			// @todo Please remove those attributes that should not be searched.
			array('PersonID, Created, CreatedBy, RoleID, UserName, FirstName, LastName, PersonEmail, Password, Gender, Enabled, LastLogin', 'safe', 'on'=>'search'),
		);
	}

	/**
	 * @return array relational rules.
	 */
	public function relations()
	{
		// NOTE: you may need to adjust the relation name and the related
		// class name for the relations automatically generated below.
		return array(
			'role' => array(self::BELONGS_TO, 'Roles', 'RoleID'),
			'createdBy' => array(self::BELONGS_TO, 'Persons', 'CreatedBy'),
			'persons' => array(self::HAS_MANY, 'Persons', 'CreatedBy'),
			'personsCities' => array(self::HAS_MANY, 'PersonsCities', 'PersonID'),
			'personsMasters' => array(self::HAS_MANY, 'PersonsMasters', 'PersonID'),
		);
	}

	/**
	 * @return array customized attribute labels (name=>label)
	 */
	public function attributeLabels()
	{
		return array(
			'PersonID' => 'Person',
			'Created' => 'Created',
			'CreatedBy' => 'Created By',
			'RoleID' => 'Role',
			'UserName' => 'User Name',
			'FirstName' => 'First Name',
			'LastName' => 'Last Name',
			'PersonEmail' => 'Person Email',
			'Password' => 'Password',
			'Gender' => 'Gender',
			'Enabled' => 'Enabled',
			'LastLogin' => 'Last Login',
		);
	}

	/**
	 * Retrieves a list of models based on the current search/filter conditions.
	 *
	 * Typical usecase:
	 * - Initialize the model fields with values from filter form.
	 * - Execute this method to get CActiveDataProvider instance which will filter
	 * models according to data in model fields.
	 * - Pass data provider to CGridView, CListView or any similar widget.
	 *
	 * @return CActiveDataProvider the data provider that can return the models
	 * based on the search/filter conditions.
	 */
	public function search()
	{
		// @todo Please modify the following code to remove attributes that should not be searched.

		$criteria=new CDbCriteria;

		$criteria->compare('PersonID',$this->PersonID,true);
		$criteria->compare('Created',$this->Created,true);
		$criteria->compare('CreatedBy',$this->CreatedBy,true);
		$criteria->compare('RoleID',$this->RoleID);
		$criteria->compare('UserName',$this->UserName,true);
		$criteria->compare('FirstName',$this->FirstName,true);
		$criteria->compare('LastName',$this->LastName,true);
		$criteria->compare('PersonEmail',$this->PersonEmail,true);
		$criteria->compare('Password',$this->Password,true);
		$criteria->compare('Gender',$this->Gender,true);
		$criteria->compare('Enabled',$this->Enabled);
		$criteria->compare('LastLogin',$this->LastLogin,true);

		return new CActiveDataProvider($this, array(
			'criteria'=>$criteria,
		));
	}

	/**
	 * Returns the static model of the specified AR class.
	 * Please note that you should have this exact method in all your CActiveRecord descendants!
	 * @param string $className active record class name.
	 * @return Persons the static model class
	 */
	public static function model($className=__CLASS__)
	{
		return parent::model($className);
	}
}
